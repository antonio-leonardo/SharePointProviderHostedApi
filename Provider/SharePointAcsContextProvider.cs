using System;
using System.Web;
using System.Net;
using Microsoft.IdentityModel.Tokens;

using SharePointProviderHostedApi.Types;
using SharePointProviderHostedApi.Token;
using SharePointProviderHostedApi.Context;

namespace SharePointProviderHostedApi.Provider
{
    /// <summary>
    /// Default provider for SharePointAcsContext.
    /// </summary>
    internal sealed class SharePointAcsContextProvider : SharePointContextProvider
    {
        protected override SharePointContext CreateSharePointContext
            (Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
        {
            string contextTokenString = ProcessTokenStrings.GetContextTokenFromRequest(httpRequest);
            if (string.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }
            SharePointContextToken contextToken = null;
            try
            {
                contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority);
            }
            catch (WebException)
            {
                return null;
            }
            catch (AudienceUriValidationFailedException)
            {
                return null;
            }
            return new SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken);
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;
            if (spAcsContext != null)
            {
                Uri spHostUrl = TokenHelper.GetSPHostUrl(httpContext.Request);
                string contextToken = ProcessTokenStrings.GetContextTokenFromRequest(httpContext.Request);
                HttpCookie spCacheKeyCookie = httpContext.Request.Cookies[SharePointKeys.SPCacheKey.ToString()];
                string spCacheKey = spCacheKeyCookie != null ? spCacheKeyCookie.Value : null;

                return spHostUrl == spAcsContext.SPHostUrl &&
                       !string.IsNullOrEmpty(spAcsContext.CacheKey) &&
                       spCacheKey == spAcsContext.CacheKey &&
                       !string.IsNullOrEmpty(spAcsContext.ContextToken) &&
                       (string.IsNullOrEmpty(contextToken) || contextToken == spAcsContext.ContextToken);
            }
            return false;
        }

        protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
        {
            return httpContext.Session[SharePointKeys.SPContext.ToString()] as SharePointAcsContext;
        }

        protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;
            if (spAcsContext != null)
            {
                HttpCookie spCacheKeyCookie = new HttpCookie(SharePointKeys.SPCacheKey.ToString())
                {
                    Value = spAcsContext.CacheKey,
                    Secure = true,
                    HttpOnly = true
                };

                httpContext.Response.AppendCookie(spCacheKeyCookie);
            }
            httpContext.Session[SharePointKeys.SPContext.ToString()] = spAcsContext;
        }
    }
}