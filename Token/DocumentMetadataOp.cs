using System;
using System.Net;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Security.Cryptography.X509Certificates;

using ZCR.SharePointFramework.CSOM.Json;
using ZCR.SharePointFramework.CSOM.Types;

namespace ZCR.SharePointFramework.CSOM.Token
{
    internal static class DocumentMetadataOp
    {
        internal static X509Certificate2 GetAcsSigningCert(string realm)
        {
            JsonMetadataDocument document = GetMetadataDocument(realm);

            if (null != document.keys && document.keys.Length > 0)
            {
                JsonKey signingKey = document.keys[0];

                if (null != signingKey && null != signingKey.keyValue)
                {
                    return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                }
            }

            throw new Exception("Metadata document does not contain ACS signing certificate.");
        }

        internal static string GetDelegationServiceUrl(string realm)
        {
            JsonMetadataDocument document = GetMetadataDocument(realm);

            JsonEndpoint delegationEndpoint = GetEndpoint(document.endpoints, ACSMetadatProtocol.DelegationIssuance).ToArray()[0];

            if (null != delegationEndpoint)
            {
                return delegationEndpoint.location;
            }
            throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
        }

        internal static string GetStsUrl(string realm)
        {
            JsonMetadataDocument document = GetMetadataDocument(realm);

            JsonEndpoint s2sEndpoint = GetEndpoint(document.endpoints, ACSMetadatProtocol.S2SProtocol).ToArray()[0];

            if (null != s2sEndpoint)
            {
                return s2sEndpoint.location;
            }

            throw new Exception("Metadata document does not contain STS endpoint url");
        }

        private static JsonMetadataDocument GetMetadataDocument(string realm)
        {
            string acsMetadataEndpointUrlWithRealm = string.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                   ProcessTokenStrings.GetAcsMetadataEndpointUrl(),
                                                                   realm);
            byte[] acsMetadata = null;
            using (WebClient webClient = new WebClient())
            {
                acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
            }
            string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

            if (null == document)
            {
                throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
            }

            return document;
        }     

        private static IEnumerable<JsonEndpoint> GetEndpoint(JsonEndpoint[] endpoints, ACSMetadatProtocol protocol)
        {
            for (int i = 0; i < endpoints.Length; i++)
            {
                if (endpoints[i].protocol == protocol.XmlEnumValue())
                {
                    yield return endpoints[i];
                    break;
                }
            }
        }
    }
}
