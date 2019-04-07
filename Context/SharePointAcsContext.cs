using System;
using System.IO;
using System.Net;
using System.ServiceModel;
using System.Security.Principal;
using Microsoft.SharePoint.Client;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.SharePoint.Client.EventReceivers;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;

using SharePointProviderHostedApi.Token;

namespace SharePointProviderHostedApi.Context
{
    /// <summary>
    /// Encapsulates all the information from SharePoint in ACS mode.
    /// </summary>
    internal sealed class SharePointAcsContext : SharePointContext
    {
        private readonly string contextToken;
        private readonly SharePointContextToken contextTokenObj;

        /// <summary>
        /// The context token.
        /// </summary>
        internal string ContextToken
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextToken : null; }
        }

        /// <summary>
        /// The context token's "CacheKey" claim.
        /// </summary>
        internal string CacheKey
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextTokenObj.CacheKey : null; }
        }

        /// <summary>
        /// The context token's "refreshtoken" claim.
        /// </summary>
        internal string RefreshToken
        {
            get { return this.contextTokenObj.ValidTo > DateTime.UtcNow ? this.contextTokenObj.RefreshToken : null; }
        }

        internal override string UserAccessTokenForSPHost
        {
            get
            {
                return this.GetAccessTokenString(ref this.userAccessTokenForSPHost, () => TokenHelper.GetAccessToken(this.contextTokenObj, this.SPHostUrl.Authority));
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
                return this.GetAccessTokenString(ref this.userAccessTokenForSPAppWeb, () => TokenHelper.GetAccessToken(this.contextTokenObj, this.SPAppWebUrl.Authority));
            }
        }

        internal override string AppOnlyAccessTokenForSPHost
        {
            get
            {
                return this.GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost, () => this.GetAppOnlyAccessToken(
                    ProcessTokenStrings.SharePointPrincipal, this.SPHostUrl.Authority, ProcessTokenStrings.GetRealmFromTargetUrl(this.SPHostUrl)));
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

                return this.GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb, () => this.GetAppOnlyAccessToken(
                    ProcessTokenStrings.SharePointPrincipal, this.SPAppWebUrl.Authority, ProcessTokenStrings.GetRealmFromTargetUrl(this.SPAppWebUrl)));
            }
        }

        internal SharePointAcsContext
            (Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, string contextToken, SharePointContextToken contextTokenObj)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (string.IsNullOrEmpty(contextToken))
            {
                throw new ArgumentNullException("contextToken");
            }

            if (contextTokenObj == null)
            {
                throw new ArgumentNullException("contextTokenObj");
            }

            this.contextToken = contextToken;
            this.contextTokenObj = contextTokenObj;
        }

        /// <summary>
        /// Creates a client context based on the properties of a remote event receiver
        /// </summary>
        /// <param name="properties">Properties of a remote event receiver</param>
        /// <returns>A ClientContext ready to call the web where the event originated</returns>
        internal ClientContext CreateRemoteEventReceiverClientContext(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl;
            if (properties.ListEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ListEventProperties.WebUrl);
            }
            else if (properties.ItemEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            }
            else if (properties.WebEventProperties != null)
            {
                sharepointUrl = new Uri(properties.WebEventProperties.FullUrl);
            }
            else
            {
                return null;
            }
            if (WebConfigAddInDataRescue.IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }
            return this.CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// Retrieves an access token from ACS using the specified authorization code, and uses that access token to 
        /// create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="authorizationCode">Authorization code to use when retrieving the access token from ACS</param>
        /// <param name="redirectUri">Redirect URI registered for this add-in</param>
        /// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
        internal ClientContext GetClientContextWithAuthorizationCode(string targetUrl, string authorizationCode, Uri redirectUri)
        {
            return this.GetClientContextWithAuthorizationCode(targetUrl, ProcessTokenStrings.SharePointPrincipal, authorizationCode, ProcessTokenStrings.GetRealmFromTargetUrl(new Uri(targetUrl)), redirectUri);
        }

        /// <summary>
        /// Retrieves an access token from ACS using the specified authorization code, and uses that access token to 
        /// create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="targetPrincipalName">Name of the target SharePoint principal</param>
        /// <param name="authorizationCode">Authorization code to use when retrieving the access token from ACS</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <param name="redirectUri">Redirect URI registered for this add-in</param>
        /// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
        private ClientContext GetClientContextWithAuthorizationCode(string targetUrl, string targetPrincipalName, string authorizationCode, string targetRealm, Uri redirectUri)
        {
            Uri targetUri = new Uri(targetUrl);
            string accessToken =
                TokenHelper.GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken;
            return this.GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// Retrieves an access token from ACS using the specified context token, and uses that access token to create
        /// a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="contextTokenString">Context token received from the target SharePoint site</param>
        /// <param name="appHostUrl">Url authority of the hosted add-in.  If this is null, the value in the HostedAppHostName
        /// of web.config will be used instead</param>
        /// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
        internal ClientContext GetClientContextWithContextToken(string targetUrl, string contextTokenString, string appHostUrl)
        {
            SharePointContextToken contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, appHostUrl);
            Uri targetUri = new Uri(targetUrl);
            string accessToken = TokenHelper.GetAccessToken(contextToken, targetUri.Authority).AccessToken;
            return this.GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// Retrieves an app-only access token from ACS to call the specified principal 
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <returns>An access token with an audience of the target principal</returns>
        internal OAuth2AccessTokenResponse GetAppOnlyAccessToken(string targetPrincipalName, string targetHost, string targetRealm)
        {
            if (targetRealm == null)
            {
                targetRealm = WebConfigAddInDataRescue.Realm;
            }

            string resource = ProcessTokenStrings.GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = ProcessTokenStrings.GetFormattedPrincipal(WebConfigAddInDataRescue.ClientId, WebConfigAddInDataRescue.HostedAppHostName, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, WebConfigAddInDataRescue.ClientSecret, resource);
            oauth2Request.Resource = resource;

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();

            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(DocumentMetadataOp.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }
            return oauth2Response;
        }

        /// <summary>
        /// Ensures the access token is valid and returns it.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        /// <returns>The access token string.</returns>
        private string GetAccessTokenString(ref Tuple<string, DateTime> accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler)
        {
            this.RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);
            return TokenHelper.IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
        }

        /// <summary>
        /// Renews the access token if it is not valid.
        /// </summary>
        /// <param name="accessToken">The access token to renew.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        private void RenewAccessTokenIfNeeded(ref Tuple<string, DateTime> accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler)
        {
            if (TokenHelper.IsAccessTokenValid(accessToken))
            {
                return;
            }
            try
            {
                OAuth2AccessTokenResponse oAuth2AccessTokenResponse = tokenRenewalHandler();
                DateTime expiresOn = oAuth2AccessTokenResponse.ExpiresOn;
                if ((expiresOn - oAuth2AccessTokenResponse.NotBefore) > AccessTokenLifetimeTolerance)
                {
                    // Make the access token get renewed a bit earlier than the time when it expires
                    // so that the calls to SharePoint with it will have enough time to complete successfully.
                    expiresOn -= AccessTokenLifetimeTolerance;
                }
                accessToken = Tuple.Create(oAuth2AccessTokenResponse.AccessToken, expiresOn);
            }
            catch (WebException)
            {
            }
        }

        /// <summary>
        /// Creates a client context based on the properties of an add-in event
        /// </summary>
        /// <param name="properties">Properties of an add-in event</param>
        /// <param name="useAppWeb">True to target the app web, false to target the host web</param>
        /// <returns>A ClientContext ready to call the app web or the parent web</returns>
        private ClientContext CreateAppEventClientContext(SPRemoteEventProperties properties, bool useAppWeb)
        {
            if (properties.AppEventProperties == null)
            {
                return null;
            }
            Uri sharepointUrl = useAppWeb ? properties.AppEventProperties.AppWebFullUrl : properties.AppEventProperties.HostWebFullUrl;
            if (WebConfigAddInDataRescue.IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }
            return this.CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// Retrieves an S2S client context with an access token signed by the application's private certificate on 
        /// behalf of the specified WindowsIdentity and intended for application at the targetApplicationUri using the 
        /// targetRealm. If no Realm is specified in web.config, an auth challenge will be issued to the 
        /// targetApplicationUri to discover it.
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <param name="identity">Windows identity of the user on whose behalf to create the access token</param>
        /// <returns>A ClientContext using an access token with an audience of the target application</returns>
        private ClientContext GetS2SClientContextWithWindowsIdentity(Uri targetApplicationUri, WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(WebConfigAddInDataRescue.Realm) ? ProcessTokenStrings.GetRealmFromTargetUrl(targetApplicationUri) : WebConfigAddInDataRescue.Realm;
            JsonWebTokenClaim[] claims = identity != null ? ProcessTokenStrings.GetClaimsWithWindowsIdentity(identity) : null;
            string accessToken = ProcessTokenStrings.GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
            return this.GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }

        private ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl)
        {
            string contextTokenString = properties.ContextToken;

            if (string.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }

            SharePointContextToken contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
            string accessToken = TokenHelper.GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;
            return this.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
        }
    }
}