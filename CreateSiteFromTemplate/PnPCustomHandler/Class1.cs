using Microsoft.SharePoint.Client;

using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Extensibility;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnPCustomHandler
{
    public class FolderItem
    {
        public string FolderName { get; set; }
        public string ParentUrl { get; set; }
    }
    public class FolderExtensibilityHandler : IProvisioningExtensibilityHandler
    {
        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope, string configurationData)
        {
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();
            var templAllLists = creationInformation.BaseTemplate.Lists;
            foreach (var templist in templAllLists)
            {
                if (templist.Url == "Shared Documents")
                {
                    var spList = ctx.Web.GetList(web.ServerRelativeUrl + "/" + templist.Url);
                    ctx.Load(spList);
                    ctx.Load(spList.RootFolder);
                    ctx.ExecuteQuery();
                    
                    
                    ListItemCollection folderItems = spList.GetItems(new CamlQuery() { ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>" });
                    ctx.Load(folderItems);
                    ctx.ExecuteQuery();
                    List<FolderItem> folders = new List<FolderItem>();
                    foreach (var folderItem in folderItems)
                    {
                        folders.Add(new FolderItem()
                        {
                            FolderName = Convert.ToString(folderItem["FileLeafRef"]),
                            ParentUrl = Convert.ToString(folderItem["FileDirRef"]).Replace(spList.RootFolder.ServerRelativeUrl, String.Empty).TrimStart(new char[] { '/' })
                        });
                    }
                    var parentFolders = folders.Where(f => string.IsNullOrEmpty(f.ParentUrl)).ToList();                    
                    foreach(var parentFolder in parentFolders)
                    {
                        templist.Folders.Add(GetPnPFolder(creationInformation.BaseTemplate, parentFolder, folders,string.Empty));
                    }


                }
            }
            return template;
        }
        private PnP.Framework.Provisioning.Model.Folder GetPnPFolder(ProvisioningTemplate template,FolderItem folderItem, List<FolderItem> allFolderItems,string parentUrl)
        {
            PnP.Framework.Provisioning.Model.Folder folder = new PnP.Framework.Provisioning.Model.Folder(folderItem.FolderName);

            parentUrl = string.IsNullOrEmpty(parentUrl) ? folderItem.FolderName : (parentUrl + "/" + folderItem.FolderName);

            var allChildFolders = allFolderItems.Where(a => a.ParentUrl == parentUrl).ToList();

            foreach(var childFolder in allChildFolders)
            {
                folder.Folders.Add(GetPnPFolder(template, childFolder, allFolderItems, parentUrl));
            }

            return folder;


        }
        public IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData)
        {
            return null;
        }

        public void Provision(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation, TokenParser tokenParser, PnPMonitoredScope scope, string configurationData)
        {

        }
    }
}
