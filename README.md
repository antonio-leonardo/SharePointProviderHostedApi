# SharePoint Provider-Hosted Api
This designed API is a result from a Visual Studio SharePoint Add-In Template Project refactoring, to be useful at SharePoint 2016 Hosted, SharePoint OnLine or Hybrid Farm; this is can a alternative if do you want a decople solution for C# Provider Hosted Solutions. This was tested in SharePoint 2016 and SharePoint Online project customizations using Asp.NET Web Forms

# Dependecies
Needed to install a nuget package for [Microsoft.Identity.Model](https://www.nuget.org/packages/Microsoft.IdentityModel/) and [Newtonsoft.Json](https://www.newtonsoft.com/json)
and/or
Microsoft SharePoint Components SDK for [SharePoint Online](https://www.microsoft.com/en-us/download/details.aspx?id=42038) and [SharePoint 2016](https://www.microsoft.com/en-us/download/details.aspx?id=51679), this second way is better.

# Examples of use
Instantiate [ClientContext](https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee538685(v%3Doffice.15)) object from this API object called [SharePointContextCSOM](https://github.com/antonio-leonardo/SharePointProviderHostedApi/blob/master/SharePointContextCSOM.cs) and all the magic is will be executed! Because the has all specified sharepoint artefact items that developer needs to contextualize, giving to developer only responsability to make C.R.U.D. roles.

----------------------
## License

[View MIT license](https://github.com/antonio-leonardo/SharePointProviderHostedApi/blob/master/LICENSE)