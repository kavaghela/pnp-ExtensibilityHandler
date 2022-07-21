using Microsoft.SharePoint.Client;

using PnP.Framework;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Extensibility;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

using PnPCustomHandler;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateSiteFromTemplate
{
    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {
                using (ClientContext sourceSiteContext = new AuthenticationManager("", @"", "", "").GetContext("https://contoso.sharepoint.com/sites/Template"))
                {
                    Web srcWeb = sourceSiteContext.Web;
                    sourceSiteContext.Load(srcWeb);
                    sourceSiteContext.ExecuteQuery();

                    Console.WriteLine(srcWeb.Title);
                    var assemblyName = typeof(FolderExtensibilityHandler).Assembly.FullName;
                    var handler = new ExtensibilityHandler()
                    {
                        Assembly = assemblyName,
                        Configuration = "",
                        Enabled = true,
                        Type = "PnPCustomHandler.FolderExtensibilityHandler",

                    };


                    var creationInformation = new ProvisioningTemplateCreationInformation(sourceSiteContext.Web)
                    {
                        ExtensibilityHandlers = new List<ExtensibilityHandler>() { handler },
                        IncludeSiteGroups = false,
                        IncludeTermGroupsSecurity = false,
                        IncludeSiteCollectionTermGroup = false,
                        IncludeSearchConfiguration = false,
                        IncludeAllClientSidePages = false,
                        ProgressDelegate = (message, step, total) =>
                        {
                            Console.WriteLine($"Template Retrieval Progress: {step}/{total} - {message}");
                        }
                    };

                    var template = srcWeb.GetProvisioningTemplate(creationInformation);

                    using (ClientContext destSiteContext = new AuthenticationManager("", @"", "", "").GetContext("https://contoso.sharepoint.com/sites/destSite"))
                    {

                        var provisioningTemplateInformation = new ProvisioningTemplateApplyingInformation()
                        {                                                      
                            MessagesDelegate = (message, ProvisioningMessageType) =>
                            {
                                Console.WriteLine($"Provisioning Messages: {message}. Type: {ProvisioningMessageType}");
                            },
                            ProgressDelegate = (message, step, total) =>
                            {
                                Console.WriteLine($"Provisioning Progress: {step}/{total} - {message}");
                            }
                        };

                        destSiteContext.Web.ApplyProvisioningTemplate(creationInformation.BaseTemplate, provisioningTemplateInformation);
                        destSiteContext.ExecuteQuery();
                    }




                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());

            }
            finally
            {
                Console.WriteLine("Finished");
                Console.ReadKey();
            }


        }
    }


}
