using System;
using System.Reflection;
using System.Xml.Serialization;

namespace ZCR.SharePointFramework.CSOM.Types
{
    internal enum ACSMetadatProtocol
    {
        [XmlEnum("OAuth2")]
        S2SProtocol,

        [XmlEnum("DelegationIssuance1.0")]
        DelegationIssuance
    }

    /// <summary>
    /// 
    /// </summary>
    internal static class Extensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string XmlEnumValue(this Enum value)
        {
            Type type = value.GetType();
            FieldInfo fieldInfo = type.GetField(value.ToString());
            // Get the stringvalue attributes  
            XmlEnumAttribute[] attribs = fieldInfo.GetCustomAttributes(
                 typeof(XmlEnumAttribute), false) as XmlEnumAttribute[];
            // Return the first if there was a match.  
            return attribs.Length > 0 ? attribs[0].Name : null;
        }
    }
}