using System;
using System.IO;
using System.Web;
using System.Net;
using System.Globalization;
using System.Security.Principal;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.IdentityModel.Tokens;
using System.IdentityModel.Selectors;
using System.Collections.ObjectModel;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction;
using SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration;

using ZCR.SharePointFramework.CSOM.Types;

namespace ZCR.SharePointFramework.CSOM.Token
{
    internal static class TokenHelper
    {
        /// <summary>
        /// Uses the specified access token to create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="accessToken">Access token to be used when calling the specified targetUrl</param>
        /// <returns>A ClientContext ready to call targetUrl with the specified access token</returns>
        internal static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);
            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        internal static ClientContext GetS2SClientContextWithWindowsIdentity(Uri targetApplicationUri, WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(WebConfigAddInDataRescue.Realm) ? ProcessTokenStrings.GetRealmFromTargetUrl(targetApplicationUri) : WebConfigAddInDataRescue.Realm;
            JsonWebTokenClaim[] claims = identity != null ? ProcessTokenStrings.GetClaimsWithWindowsIdentity(identity) : null;
            string accessToken = ProcessTokenStrings.GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
            return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }

        private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler()
        {
            JsonWebSecurityTokenHandler handler = new JsonWebSecurityTokenHandler();
            handler.Configuration = new SecurityTokenHandlerConfiguration();
            handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
            handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            List<byte[]> securityKeys = new List<byte[]>();
            securityKeys.Add(Convert.FromBase64String(WebConfigAddInDataRescue.ClientSecret));
            if (!string.IsNullOrEmpty(WebConfigAddInDataRescue.SecondaryClientSecret))
            {
                securityKeys.Add(Convert.FromBase64String(WebConfigAddInDataRescue.SecondaryClientSecret));
            }

            List<SecurityToken> securityTokens = new List<SecurityToken>();
            securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            handler.Configuration.IssuerTokenResolver =
                SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                new ReadOnlyCollection<SecurityToken>(securityTokens),
                false);
            SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            foreach (byte[] securitykey in securityKeys)
            {
                issuerNameRegistry.AddTrustedIssuer(securitykey, ProcessTokenStrings.GetAcsPrincipalName(WebConfigAddInDataRescue.ServiceNamespace));
            }
            handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        /// <summary>
        /// Determines if the specified access token is valid.
        /// It considers an access token as not valid if it is null, or it has expired.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <returns>True if the access token is valid.</returns>
        internal static bool IsAccessTokenValid(Tuple<string, DateTime> accessToken)
        {
            return accessToken != null &&
                   !string.IsNullOrEmpty(accessToken.Item1) &&
                   accessToken.Item2 > DateTime.UtcNow;
        }

        /// <summary>
        /// Gets the SharePoint host url from QueryString of the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The specified HTTP request.</param>
        /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        internal static Uri GetSPHostUrl(HttpRequest httpRequest)
        {
            return GetSPHostUrl(new HttpRequestWrapper(httpRequest));
        }

        /// <summary>
        /// Gets the SharePoint host url from QueryString of the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The specified HTTP request.</param>
        /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        internal static Uri GetSPHostUrl(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
            {
                throw new ArgumentNullException("httpRequest");
            }
            string spHostUrlString = ProcessTokenStrings.EnsureTrailingSlash(httpRequest.QueryString[SharePointKeys.SPHostUrl.ToString()]);
            Uri spHostUrl;
            if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
            {
                return spHostUrl;
            }
            return null;
        }

        /// <summary>
        /// Validate that a specified context token string is intended for this application based on the parameters 
        /// specified in web.config. Parameters used from web.config used for validation include ClientId, 
        /// HostedAppHostNameOverride, HostedAppHostName, ClientSecret, and Realm (if it is specified). If HostedAppHostNameOverride is present,
        /// it will be used for validation. Otherwise, if the <paramref name="appHostName"/> is not 
        /// null, it is used for validation instead of the web.config's HostedAppHostName. If the token is invalid, an 
        /// exception is thrown. If the token is valid, TokenHelper's static STS metadata url is updated based on the token contents
        /// and a JsonWebSecurityToken based on the context token is returned.
        /// </summary>
        /// <param name="contextTokenString">The context token to validate</param>
        /// <param name="appHostName">The URL authority, consisting of  Domain Name System (DNS) host name or IP address and the port number, to use for token audience validation.
        /// If null, HostedAppHostName web.config setting is used instead. HostedAppHostNameOverride web.config setting, if present, will be used 
        /// for validation instead of <paramref name="appHostName"/> .</param>
        /// <returns>A JsonWebSecurityToken based on the context token.</returns>
        internal static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            JsonWebSecurityTokenHandler tokenHandler = CreateJsonWebSecurityTokenHandler();
            SecurityToken securityToken = tokenHandler.ReadToken(contextTokenString);
            JsonWebSecurityToken jsonToken = securityToken as JsonWebSecurityToken;
            SharePointContextToken token = SharePointContextToken.Create(jsonToken);

            //Alteração para HighTrust: Comentada as linhas abaixo
            //string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            //int firstDot = stsAuthority.IndexOf('.');

            //ProcessTokenStrings.GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            //ProcessTokenStrings.AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            tokenHandler.ValidateToken(jsonToken);

            string[] acceptableAudiences;
            if (!string.IsNullOrEmpty(WebConfigAddInDataRescue.HostedAppHostNameOverride))
            {
                acceptableAudiences = WebConfigAddInDataRescue.HostedAppHostNameOverride.Split(';');
            }
            else if (appHostName == null)
            {
                acceptableAudiences = new[] { WebConfigAddInDataRescue.HostedAppHostName };
            }
            else
            {
                acceptableAudiences = new[] { appHostName };
            }

            bool validationSuccessful = false;
            string realm = WebConfigAddInDataRescue.Realm ?? token.Realm;
            foreach (string audience in acceptableAudiences)
            {
                string principal = ProcessTokenStrings.GetFormattedPrincipal(WebConfigAddInDataRescue.ClientId, audience, realm);
                if (StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal))
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
            {
                throw new AudienceUriValidationFailedException(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", String.Join(";", acceptableAudiences), token.Audience));
            }

            return token;
        }

        /// <summary>
        /// Retrieves an access token from ACS to call the source of the specified context token at the specified 
        /// targetHost. The targetHost must be registered for the principal that sent the context token.
        /// </summary>
        /// <param name="contextToken">Context token issued by the intended access token audience</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <returns>An access token with an audience matching the context token's source</returns>
        internal static OAuth2AccessTokenResponse GetAccessToken(SharePointContextToken contextToken, string targetHost)
        {
            string targetPrincipalName = contextToken.TargetPrincipalName;
            // Extract the refreshToken from the context token
            string refreshToken = contextToken.RefreshToken;

            if (string.IsNullOrEmpty(refreshToken))
            {
                return null;
            }
            string targetRealm = WebConfigAddInDataRescue.Realm ?? contextToken.Realm;
            return GetAccessToken(refreshToken, targetPrincipalName, targetHost, targetRealm);
        }

        /// <summary>
        /// Uses the specified authorization code to retrieve an access token from ACS to call the specified principal 
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="authorizationCode">Authorization code to exchange for access token</param>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <param name="redirectUri">Redirect URI registered for this add-in</param>
        /// <returns>An access token with an audience of the target principal</returns>
        internal static OAuth2AccessTokenResponse GetAccessToken(string authorizationCode, string targetPrincipalName, string targetHost, string targetRealm, Uri redirectUri)
        {
            if (targetRealm == null)
            {
                targetRealm = WebConfigAddInDataRescue.Realm;
            }
            string resource = ProcessTokenStrings.GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = ProcessTokenStrings.GetFormattedPrincipal(WebConfigAddInDataRescue.ClientId, null, targetRealm);

            // Create request for token. The RedirectUri is null here.  This will fail if redirect uri is registered
            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory
                .CreateAccessTokenRequestWithAuthorizationCode(clientId, WebConfigAddInDataRescue.ClientSecret, authorizationCode, redirectUri, resource);

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response = client.Issue(DocumentMetadataOp.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
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
        /// Uses the specified refresh token to retrieve an access token from ACS to call the specified principal 
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="refreshToken">Refresh token to exchange for access token</param>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <returns>An access token with an audience of the target principal</returns>
        internal static OAuth2AccessTokenResponse GetAccessToken(string refreshToken, string targetPrincipalName, string targetHost, string targetRealm)
        {
            if (targetRealm == null)
            {
                targetRealm = WebConfigAddInDataRescue.Realm;
            }

            string resource = ProcessTokenStrings.GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = ProcessTokenStrings.GetFormattedPrincipal(WebConfigAddInDataRescue.ClientId, null, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory
                .CreateAccessTokenRequestWithRefreshToken(clientId, WebConfigAddInDataRescue.ClientSecret, refreshToken, resource);
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
    }
}