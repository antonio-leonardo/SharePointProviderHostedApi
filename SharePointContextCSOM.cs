using System;
using System.Web;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

using ZCR.SharePointFramework.CSOM.Token;
using ZCR.SharePointFramework.CSOM.Context;
using ZCR.SharePointFramework.CSOM.Provider;

namespace ZCR.SharePointFramework.CSOM
{
    public enum ContextOption
    {
        OneListGuid,
        MultipleListsGuids,
        OneListName,
        MultipleListsNames
    }

    /// <summary>
    /// Contexto SharePoint para Codificação Server-side
    /// </summary>
    /// <typeparam name="TEntity">Classe de Entidade que contenha as definições de Atributos para estabelecimento do contrato com o Back-End SharePoint</typeparam>
    public class SharePointContextCSOM<TEntity> : IDisposable
    {
        #region ' Declarações e Propriedades da Classe '

        private Guid ListGuid { get; set; }

        private string ListName { get; set; }

        private string[] ListNames { get; set; }

        private Guid[] ListGuids { get; set; }

        private List ListContext { get; set; }

        private Dictionary<string, List> ListsContext { get; set; }

        public ClientContext SharePointClientCtx { get; private set; }

        #endregion

        #region ' Construtores da Classe '

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="listName">Nome da Lista</param>
        public SharePointContextCSOM(HttpContext httpContext, string listName)
        {
            this.ListName = listName;
            this.ExecuteConstructor(httpContext, ContextOption.OneListName);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="listGUID">ID da Lista</param>
        public SharePointContextCSOM(HttpContext httpContext, Guid listGUID)
        {
            this.ListGuid = listGUID;
            this.ExecuteConstructor(httpContext, ContextOption.OneListGuid);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="siteCollectionUrl">URL da Coleção de Site para acesso Cross-Site</param>
        /// <param name="listName">Nome da Lista</param>
        public SharePointContextCSOM(HttpContext httpContext, string siteCollectionUrl, string listName)
        {
            this.ListName = listName;
            this.ExecuteConstructor(httpContext, ContextOption.OneListName, siteCollectionUrl);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="siteCollectionUrl">URL da Coleção de Site para acesso Cross-Site</param>
        /// <param name="listGUID">ID da Lista</param>
        public SharePointContextCSOM(HttpContext httpContext, string siteCollectionUrl, Guid listGUID)
        {
            this.ListGuid = listGUID;
            this.ExecuteConstructor(httpContext, ContextOption.OneListGuid, siteCollectionUrl);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="listNames">Nomes das Listas</param>
        public SharePointContextCSOM(HttpContext httpContext, params string[] listNames)
        {
            this.ListNames = new string[] { };
            this.ListNames = listNames;
            this.ExecuteConstructor(httpContext, ContextOption.MultipleListsNames);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="listNames">Nomes das Listas</param>
        public SharePointContextCSOM(HttpContext httpContext, params Guid[] listGUIDs)
        {
            this.ListGuids = new Guid[] { };
            this.ListGuids = ListGuids;
            this.ExecuteConstructor(httpContext, ContextOption.MultipleListsGuids);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="siteCollectionUrl">URL da Coleção de Site para acesso Cross-Site</param>
        /// <param name="listNames">Nomes das Listas</param>
        public SharePointContextCSOM(HttpContext httpContext, string siteCollectionUrl, params string[] listNames)
        {
            this.ListNames = new string[] { };
            this.ListNames = listNames;
            this.ExecuteConstructor(httpContext, ContextOption.MultipleListsNames, siteCollectionUrl);
        }

        /// <summary>
        /// Construtor com para contextualizar o SharePoint para Add-In ou Apps,
        /// dando suporte a autenticação Low-trut ou Hight-trust
        /// </summary>
        /// <param name="context">Contexto HTTP</param>
        /// <param name="siteCollectionUrl">URL da Coleção de Site para acesso Cross-Site</param>
        /// <param name="listNames">Nomes das Listas</param>
        public SharePointContextCSOM(HttpContext httpContext, string siteCollectionUrl, params Guid[] listGUIDs)
        {
            this.ListGuids = new Guid[] { };
            this.ListGuids = ListGuids;
            this.ExecuteConstructor(httpContext, ContextOption.MultipleListsGuids, siteCollectionUrl);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="httpContext"></param>
        /// <param name="option"></param>
        /// <param name="siteCollectionUrl"></param>
        private void ExecuteConstructor(HttpContext httpContext, ContextOption option, string siteCollectionUrl = null)
        {
            this.SharePointClientCtx = null;
            try
            {
                switch (option)
                {
                    case ContextOption.OneListGuid:
                        if (!string.IsNullOrWhiteSpace(siteCollectionUrl))
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext, siteCollectionUrl);
                            this.ListContext = this.SharePointClientCtx.Web.Lists.GetById(this.ListGuid);
                        }
                        else
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext);
                            this.ListContext = this.SharePointClientCtx.Web.Lists.GetById(this.ListGuid);
                        }
                        break;
                    case ContextOption.MultipleListsGuids:
                        if (!string.IsNullOrWhiteSpace(siteCollectionUrl))
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext, siteCollectionUrl);
                            this.ListsContext = new Dictionary<string, List>();
                            this.ListsContext = this.ContextingLists(this.SharePointClientCtx.Web, this.ListGuids).ToDictionary(x => x.Key, x => x.Value);
                        }
                        else
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext);
                            this.ListsContext = new Dictionary<string, List>();
                            this.ListsContext = this.ContextingLists(this.SharePointClientCtx.Web, this.ListGuids).ToDictionary(x => x.Key, x => x.Value);
                        }
                        break;
                    case ContextOption.OneListName:
                        if (!string.IsNullOrWhiteSpace(siteCollectionUrl))
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext, siteCollectionUrl, this.ListName);
                            this.ListContext = this.SharePointClientCtx.Web.Lists.GetByTitle(this.ListName);
                        }
                        else
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext);
                            this.ListContext = this.SharePointClientCtx.Web.Lists.GetByTitle(this.ListName);
                        }
                        break;
                    case ContextOption.MultipleListsNames:
                        if (!string.IsNullOrWhiteSpace(siteCollectionUrl))
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext, siteCollectionUrl);
                            this.ListsContext = new Dictionary<string, List>();
                            this.ListsContext = this.ContextingLists(this.SharePointClientCtx.Web, this.ListNames).ToDictionary(x => x.Key, x => x.Value);
                        }
                        else
                        {
                            this.SharePointClientCtx = this.StartObjectVariables(httpContext);
                            this.ListsContext = new Dictionary<string, List>();
                            this.ListsContext = this.ContextingLists(this.SharePointClientCtx.Web, this.ListNames).ToDictionary(x => x.Key, x => x.Value);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ocorreu algum durante a rotina de autenticação ao SharePoint.", ex);
            }
        }

        /// <summary>
        /// Define o Contexto Sharepoint baseado no tipo de autenticação realizado
        /// </summary>
        /// <param name="context">Contexto Http obtido</param>
        private ClientContext StartObjectVariables(HttpContext httpContext, string siteCollectionUrl = null, string listName = null)
        {
            SharePointContextProvider ShpProvider = null;
            SharePointContextToken ContextToken = null;
            SharePointContext ShpContext = null;
            Uri SharePointUrl = null;
            ClientContext _currentContext = null;
            string AccessToken = null, contextTokenString = null;

            if (!WebConfigAddInDataRescue.IsHighTrustApp())
            {
                ShpProvider = new SharePointAcsContextProvider();
                ShpContext = ShpProvider.GetSharePointContext(httpContext);
                contextTokenString = ((SharePointAcsContext)ShpContext).ContextToken;
                ContextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpContext.Request.Url.Authority) ?? TokenHelper.ReadAndValidateContextToken(contextTokenString, SharePointUrl.Authority);
            }
            else
            {
                ShpProvider = new SharePointHighTrustContextProvider();
                ShpContext = ShpProvider.GetSharePointContext(httpContext);
            }

            if (!WebConfigAddInDataRescue.IsHighTrustApp() && string.IsNullOrWhiteSpace(siteCollectionUrl))
            {
                AccessToken = TokenHelper.GetAccessToken(ContextToken, SharePointUrl.Authority).AccessToken;
                _currentContext = ShpContext.GetClientContextWithAccessToken(SharePointUrl.ToString(), AccessToken);
            }
            else if (WebConfigAddInDataRescue.IsHighTrustApp() && string.IsNullOrWhiteSpace(siteCollectionUrl))
            {
                AccessToken = ShpContext.UserAccessTokenForSPAppWeb ?? ShpContext.UserAccessTokenForSPHost;
                _currentContext = ShpContext.CreateAppOnlyClientContextForSPHost();
            }
            else if (!string.IsNullOrWhiteSpace(siteCollectionUrl))
            {
                AccessToken = ShpContext.UserAccessTokenForSPHost;
                _currentContext = TokenHelper.GetClientContextWithAccessToken(siteCollectionUrl, AccessToken);
            }

            return _currentContext;
        }

        /// <summary>
        /// Obtém todas as Listas baseado nos GUIDs das Listas
        /// </summary>
        /// <param name="Web">Objeto SPWeb do site SharePoint</param>
        /// <param name="listGUIDs">GUIDs das Listas a sewrem contextualizadas</param>
        /// <returns>System.Collections.Generic.IEnumerable of KeyValuePair</returns>
        private IEnumerable<KeyValuePair<string, List>> ContextingLists(Web Web, params Guid[] listGUIDs)
        {
            for (int i = 0; i < listGUIDs.Length; i++)
            {
                List list = Web.Lists.GetById(listGUIDs[i]);
                if (null != list)
                {
                    yield return new KeyValuePair<string, List>(list.Title, list);
                }
                else
                {
                    throw new Exception("O Guid '" + listGUIDs[i] + "' não reflete uma Lista existente no Site SharePoint.");
                }
            }
        }

        /// <summary>
        /// Obtém todas as Listas baseado nos nomes das Listas
        /// </summary>
        /// <param name="Web">Objeto SPWeb do site SharePoint</param>
        /// <param name="listNames">Nomes das Listas a sewrem contextualizadas</param>
        /// <returns>System.Collections.Generic.IEnumerable of KeyValuePair</returns>
        private IEnumerable<KeyValuePair<string, List>> ContextingLists(Web Web, params string[] listNames)
        {
            for (int i = 0; i < listNames.Length; i++)
            {
                List list = Web.Lists.GetByTitle(listNames[i]);
                if (null != list)
                {
                    yield return new KeyValuePair<string, List>(listNames[i], list);
                }
                else
                {
                    throw new Exception("O nome '" + listNames[i] + "' não reflete uma Lista existente no Site SharePoint.");
                }
            }
        }

        #endregion

        #region ' Garbage collector '

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing && null != this.SharePointClientCtx)
            {
                this.SharePointClientCtx.Dispose();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        ~SharePointContextCSOM()
        {
            Dispose(false);
        }

        #endregion
    }
}