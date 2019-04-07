using System;
using System.IO;
using System.Net;
using System.Web;
using System.Globalization;
using System.Security.Principal;
using System.Collections.Generic;
using Microsoft.IdentityModel.S2S.Tokens;

using SharePointProviderHostedApi.Types;

namespace SharePointProviderHostedApi.Token
{
    internal static class ProcessTokenStrings
    {
        internal static string GlobalEndPointPrefix = "accounts";
        internal static string AcsHostUrl = "accesscontrol.windows.net";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";

        /// <summary>
        /// Retrieves the context token string from the specified request by looking for well-known parameter names in the 
        /// POSTed form parameters and the querystring. Returns null if no context token is found.
        /// </summary>
        /// <param name="request">HttpRequest in which to look for a context token</param>
        /// <returns>The context token string</returns>
        internal static string GetContextTokenFromRequest(HttpRequest request)
        {
            return GetContextTokenFromRequest(new HttpRequestWrapper(request));
        }

        /// <summary>
        /// Retrieves the context token string from the specified request by looking for well-known parameter names in the 
        /// POSTed form parameters and the querystring. Returns null if no context token is found.
        /// </summary>
        /// <param name="request">HttpRequest in which to look for a context token</param>
        /// <returns>The context token string</returns>
        internal static string GetContextTokenFromRequest(HttpRequestBase request)
        {
            foreach (string paramName in Enum.GetNames(typeof(RequestParams)))
            {
                if (!string.IsNullOrEmpty(request.Form[paramName]))
                {
                    return request.Form[paramName];
                }
                if (!string.IsNullOrEmpty(request.QueryString[paramName]))
                {
                    return request.QueryString[paramName];
                }
            }
            return null;
        }

        /// <summary>
        /// SharePoint principal.
        /// </summary>
        internal const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// Lifetime of HighTrust access token, 12 hours.
        /// </summary>
        internal static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        /// <summary>
        /// Retrieves an S2S access token signed by the application's private certificate on behalf of the specified 
        /// WindowsIdentity and intended for the SharePoint at the targetApplicationUri. If no Realm is specified in 
        /// web.config, an auth challenge will be issued to the targetApplicationUri to discover it.
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <param name="identity">Windows identity of the user on whose behalf to create the access token</param>
        /// <returns>An access token with an audience of the target principal</returns>
        internal static string GetS2SAccessTokenWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(WebConfigAddInDataRescue.Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : WebConfigAddInDataRescue.Realm;
            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;
            return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
        }

        internal static string GetS2SAccessTokenWithClaims(
            string targetApplicationHostName,
            string targetRealm,
            IEnumerable<JsonWebTokenClaim> claims)
        {
            return IssueToken
                (WebConfigAddInDataRescue.ClientId, WebConfigAddInDataRescue.IssuerId, targetRealm,
                SharePointPrincipal, targetRealm, targetApplicationHostName, true, claims, claims == null);
        }


        private static string IssueToken
            (string sourceApplication, string issuerApplication, string sourceRealm, string targetApplication, string targetRealm,
            string targetApplicationHostName, bool trustedForDelegation, IEnumerable<JsonWebTokenClaim> claims, bool appOnly = false)
        {
            if (null == WebConfigAddInDataRescue.SigningCredentials)
            {
                throw new InvalidOperationException("SigningCredentials was not initialized");
            }

            #region Actor token

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<JsonWebTokenClaim> actorClaims = new List<JsonWebTokenClaim>();
            actorClaims.Add(new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid));
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new JsonWebTokenClaim(ClaimType.trustedfordelegation.ToString(), "true"));
            }

            // Create token
            JsonWebSecurityToken actorToken = new JsonWebSecurityToken(
                issuer: issuer,
                audience: audience,
                validFrom: DateTime.UtcNow,
                validTo: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                signingCredentials: WebConfigAddInDataRescue.SigningCredentials,
                claims: actorClaims);

            string actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);
            
            if (appOnly)
            {
                actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);
                // App-only token is the same as actor token for delegated case
                return actorTokenString;
            }

            #endregion Actor token

            #region Outer token

            List<JsonWebTokenClaim> outerClaims = null == claims ? new List<JsonWebTokenClaim>() : new List<JsonWebTokenClaim>(claims);
            outerClaims.Add(new JsonWebTokenClaim(ClaimType.actortoken.ToString(), actorTokenString));

            JsonWebSecurityToken jsonToken = new JsonWebSecurityToken(
                nameid, // outer token issuer should match actor token nameid
                audience,
                DateTime.UtcNow,
                DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                outerClaims);

            string accessToken = new JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken);

            #endregion Outer token

            return accessToken;
        }

        internal static JsonWebTokenClaim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity)
        {
            JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
            {
                new JsonWebTokenClaim(ClaimType.nameid.ToString(), identity.User.Value.ToLower()),
                new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
            };
            return claims;
        }

        /// <summary>
        /// Get authentication realm from SharePoint
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <returns>String representation of the realm GUID</returns>
        internal static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }

        internal static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!string.IsNullOrEmpty(hostName))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return string.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        internal static string GetFormattedPrincipal(string principalName, string hostName, string realm, string globalEndPointPrefix, string acsHostUrl)
        {
            if (!string.IsNullOrEmpty(hostName))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return string.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        internal static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        internal static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return string.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
        }

        /// <summary>
        /// Returns the SharePoint url to which the add-in should redirect the browser to request consent and get back
        /// an authorization code.
        /// </summary>
        /// <param name="contextUrl">Absolute Url of the SharePoint site</param>
        /// <param name="scope">Space-delimited permissions to request from the SharePoint site in "shorthand" format 
        /// (e.g. "Web.Read Site.Write")</param>
        /// <returns>Url of the SharePoint site's OAuth authorization page</returns>
        internal static string GetAuthorizationUrl(string contextUrl, string scope)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                WebConfigAddInDataRescue.ClientId,
                scope);
        }

        /// <summary>
        /// Returns the SharePoint url to which the add-in should redirect the browser to request consent and get back
        /// an authorization code.
        /// </summary>
        /// <param name="contextUrl">Absolute Url of the SharePoint site</param>
        /// <param name="scope">Space-delimited permissions to request from the SharePoint site in "shorthand" format
        /// (e.g. "Web.Read Site.Write")</param>
        /// <param name="redirectUri">Uri to which SharePoint should redirect the browser to after consent is 
        /// granted</param>
        /// <returns>Url of the SharePoint site's OAuth authorization page</returns>
        internal static string GetAuthorizationUrl(string contextUrl, string scope, string redirectUri)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                WebConfigAddInDataRescue.ClientId,
                scope,
                redirectUri);
        }

        /// <summary>
        /// Returns the SharePoint url to which the add-in should redirect the browser to request a new context token.
        /// </summary>
        /// <param name="contextUrl">Absolute Url of the SharePoint site</param>
        /// <param name="redirectUri">Uri to which SharePoint should redirect the browser to with a context token</param>
        /// <returns>Url of the SharePoint site's context token redirect page</returns>
        internal static string GetAppContextTokenRequestUrl(string contextUrl, string redirectUri)
        {
            return string.Format(
                "{0}{1}?client_id={2}&redirect_uri={3}",
                EnsureTrailingSlash(contextUrl),
                RedirectPage,
                WebConfigAddInDataRescue.ClientId,
                redirectUri);
        }

        /// <summary>
        /// Ensures that the specified URL ends with '/' if it is not null or empty.
        /// </summary>
        /// <param name="url">The url.</param>
        /// <returns>The url ending with '/' if it is not null or empty.</returns>
        internal static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
            {
                return url + "/";
            }
            return url;
        }
    }
}