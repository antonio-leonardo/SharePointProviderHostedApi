using System;
using System.Security.Principal;

using SharePointProviderHostedApi.Token;

namespace SharePointProviderHostedApi.Context
{
    /// <summary>
    /// Encapsulates all the information from SharePoint in HighTrust mode.
    /// </summary>
    internal sealed class SharePointHighTrustContext : SharePointContext
    {
        /// <summary>
        /// The Windows identity for the current user.
        /// </summary>
        internal WindowsIdentity LogonUserIdentity { get; private set; }

        internal override string UserAccessTokenForSPHost
        {
            get
            {
                return this.GetAccessTokenString(ref this.userAccessTokenForSPHost,
                                            () => ProcessTokenStrings.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, this.LogonUserIdentity));
            }
        }

        internal override string UserAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }
                return this.GetAccessTokenString(ref this.userAccessTokenForSPAppWeb,
                                            () => ProcessTokenStrings.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, this.LogonUserIdentity));
            }
        }

        internal override string AppOnlyAccessTokenForSPHost
        {
            get
            {
                return this.GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost,
                                            () => ProcessTokenStrings.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, null));
            }
        }

        internal override string AppOnlyAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }
                return this.GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb,
                                            () => ProcessTokenStrings.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, null));
            }
        }

        internal SharePointHighTrustContext
            (Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, WindowsIdentity logonUserIdentity)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (logonUserIdentity == null)
            {
                throw new ArgumentNullException("logonUserIdentity");
            }
            this.LogonUserIdentity = logonUserIdentity;
        }

        /// <summary>
        /// Ensures the access token is valid and returns it.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        /// <returns>The access token string.</returns>
        private string GetAccessTokenString(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
        {
            this.RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);
            return TokenHelper.IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
        }

        /// <summary>
        /// Renews the access token if it is not valid.
        /// </summary>
        /// <param name="accessToken">The access token to renew.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        private void RenewAccessTokenIfNeeded(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
        {
            if (TokenHelper.IsAccessTokenValid(accessToken))
            {
                return;
            }
            DateTime expiresOn = DateTime.UtcNow.Add(ProcessTokenStrings.HighTrustAccessTokenLifetime);
            if (ProcessTokenStrings.HighTrustAccessTokenLifetime > AccessTokenLifetimeTolerance)
            {
                // Make the access token get renewed a bit earlier than the time when it expires
                // so that the calls to SharePoint with it will have enough time to complete successfully.
                expiresOn -= AccessTokenLifetimeTolerance;
            }
            accessToken = Tuple.Create(tokenRenewalHandler(), expiresOn);
        }
    }
}