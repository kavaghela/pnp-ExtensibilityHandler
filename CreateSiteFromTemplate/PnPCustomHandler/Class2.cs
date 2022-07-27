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
    
    public class PageMetadataExtensibilityHandler : IProvisioningExtensibilityHandler
    {
        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope, string configurationData)
        {
            Web web = ctx.Web;
            ctx.Load(web);
            ctx.ExecuteQuery();

            var allPages = template.ClientSidePages;
            var sitePages = web.GetSitePagesLibrary();
            foreach (var page in allPages)
            {
                var spPage =  web.LoadClientSidePage(page.PageName);
                ListItem item = sitePages.GetItemById(spPage.PageListItem.Id);
                ctx.Load(item);                
                ctx.ExecuteQuery();

                // Build Your Logic Here to add more field value 
                // As Sample I have added only Title field value
                page.FieldValues.Add("Title", Convert.ToString(item["Title"]));
                

            }
            
            return template;
        }
        private PnP.Framework.Provisioning.Model.Folder GetPnPFolder(ProvisioningTemplate template, FolderItem folderItem, List<FolderItem> allFolderItems, string parentUrl)
        {
            PnP.Framework.Provisioning.Model.Folder folder = new PnP.Framework.Provisioning.Model.Folder(folderItem.FolderName);

            parentUrl = string.IsNullOrEmpty(parentUrl) ? folderItem.FolderName : (parentUrl + "/" + folderItem.FolderName);

            var allChildFolders = allFolderItems.Where(a => a.ParentUrl == parentUrl).ToList();

            foreach (var childFolder in allChildFolders)
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
