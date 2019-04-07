using System.Web.Configuration;
using System.IdentityModel.Tokens;
using System.Security.Cryptography.X509Certificates;
using X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials;

namespace SharePointProviderHostedApi.Token
{
    internal static class WebConfigAddInDataRescue
    {
        private static readonly string _clientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId"))
            ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");

        private static readonly string _issuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId"))
            ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");

        private static readonly string _hostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");

        private static readonly string _hostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");

        private static readonly string _clientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret"))
            ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");

        private static readonly string _secondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");

        private static readonly string _realm = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string _serviceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string _clientSigningCertificatePath = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");

        private static readonly string _clientSigningCertificatePassword = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");

        private static readonly X509Certificate2 _clientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword))
            ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
        private static readonly X509SigningCredentials _signingCredentials = (_clientCertificate == null) ? null : new X509SigningCredentials(_clientCertificate, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest);


        internal static string ClientId { get { return _clientId; } }

        internal static string IssuerId { get { return _issuerId; } }

        internal static string HostedAppHostNameOverride { get { return _hostedAppHostNameOverride; } }

        internal static string HostedAppHostName { get { return _hostedAppHostName; } }

        internal static string ClientSecret { get { return _clientSecret; } }

        internal static string SecondaryClientSecret { get { return _secondaryClientSecret; } }

        internal static string Realm { get { return _realm; } }

        internal static string ServiceNamespace { get { return _serviceNamespace; } }

        internal static string ClientSigningCertificatePath { get { return _clientSigningCertificatePath; } }

        internal static string ClientSigningCertificatePassword { get { return _clientSigningCertificatePassword; } }

        internal static X509Certificate2 ClientCertificate { get { return _clientCertificate; } }

        internal static X509SigningCredentials SigningCredentials { get { return _signingCredentials; } }

        /// <summary>
        /// Determines if this is a high trust add-in.
        /// </summary>
        /// <returns>True if this is a high trust add-in.</returns>
        internal static bool IsHighTrustApp()
        {
            return SigningCredentials != null;
        }
    }
}