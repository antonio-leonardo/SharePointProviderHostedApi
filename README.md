# SharePoint Provider-Hosted Api
This designed API is a result from a Visual Studio SharePoint Add-In Template Project refactoring, to be useful at [Hight Trust](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/create-high-trust-sharepoint-add-ins) (SharePoint 2016 Hosted) or [Low Trust](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/creating-sharepoint-add-ins-that-use-low-trust-authorization) (SharePoint OnLine) or Hybrid Farm scenarios; this is can a alternative if do you want a decople solution for C# Provider Hosted Solutions. This was tested in SharePoint 2016 and SharePoint Online project customizations using Asp.NET Web Forms.

# Dependecies
  i. Needed to install a nuget package for [Microsoft.Identity.Model](https://www.nuget.org/packages/Microsoft.IdentityModel/) and [Newtonsoft.Json](https://www.newtonsoft.com/json);
  
  or
  
  ii. Microsoft SharePoint Components SDK for [SharePoint Online](https://www.microsoft.com/en-us/download/details.aspx?id=42038) and [SharePoint 2016](https://www.microsoft.com/en-us/download/details.aspx?id=51679);
  
  -The second way is better!

# Examples of use
Instantiate [ClientContext](https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee538685(v%3Doffice.15)) object from this API object called [SharePointContextCSOM](https://github.com/antonio-leonardo/SharePointProviderHostedApi/blob/master/SharePointContextCSOM.cs) and all the magic is will be succed! Because this Api has all specified sharepoint artefact items that the application needs to contextualize, giving to developer only responsability to make C.R.U.D. roles, view the example:

```cs
//For current site:
SharePointContextCSOM _providerHostedApi = new SharePointContextCSOM(HttpContext.Current, "NameOfList");

//For remote site:
SharePointContextCSOM _providerHostedApi =
              new SharePointContextCSOM(HttpContext.Current, "http://remote/site/access", "NameOfList");

using(ClientContext ctx = _providerHostedApi.SharePointClientCtx)
{
  //To do SharePoint consume business rule...
}

```
----------------------
## License

[View MIT license](https://github.com/antonio-leonardo/SharePointProviderHostedApi/blob/master/LICENSE)
