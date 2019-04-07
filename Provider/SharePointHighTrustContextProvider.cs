using System;
using System.Web;
using System.Security.Principal;

using SharePointProviderHostedApi.Types;
using SharePointProviderHostedApi.Token;
using SharePointProviderHostedApi.Context;

namespace SharePointProviderHostedApi.Provider
{
    /// <summary>
    /// Default provider for SharePointHighTrustContext.
    /// </summary>
    internal sealed class SharePointHighTrustContextProvider : SharePointContextProvider
    {
        protected override SharePointContext CreateSharePointContext
            (Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
        {
            WindowsIdentity logonUserIdentity = httpRequest.LogonUserIdentity;
            if (logonUserIdentity == null || !logonUserIdentity.IsAuthenticated || logonUserIdentity.IsGuest || logonUserIdentity.User == null)
            {
                return null;
            }

            return new SharePointHighTrustContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity);
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointHighTrustContext spHighTrustContext = spContext as SharePointHighTrustContext;

            if (spHighTrustContext != null)
            {
                Uri spHostUrl = TokenHelper.GetSPHostUrl(httpContext.Request);
                WindowsIdentity logonUserIdentity = httpContext.Request.LogonUserIdentity;

                return spHostUrl == spHighTrustContext.SPHostUrl &&
                       logonUserIdentity != null &&
                       logonUserIdentity.IsAuthenticated &&
                       !logonUserIdentity.IsGuest &&
                       logonUserIdentity.User == spHighTrustContext.LogonUserIdentity.User;
            }

            return false;
        }

        protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
        {
            return httpContext.Session[SharePointKeys.SPContext.ToString()] as SharePointHighTrustContext;
        }

        protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            httpContext.Session[SharePointKeys.SPContext.ToString()] = spContext as SharePointHighTrustContext;
        }
    }
}