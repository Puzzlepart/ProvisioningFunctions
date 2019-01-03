using System;
using System.Diagnostics;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class ApplyTemplate
    {
        //static ApplyTemplate()
        //{
        //    RedirectAssembly();
        //}

        [FunctionName("ApplyTemplate")]
        [ResponseType(typeof(ApplyTemplateResponse))]
        [Display(Name = "Apply PnP template to site", Description = "Apply a PnP template to the site.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]ApplyTemplateRequest request, TraceWriter log)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            string siteUrl = request.SiteURL;
            RedirectAssembly();
            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.TemplateURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "TemplateURL");
                }

                string templateUrl = request.TemplateURL.Trim(); // remove potential spaces/line breaks
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                var web = clientContext.Web;
                web.Lists.EnsureSiteAssetsLibrary();
                clientContext.ExecuteQueryRetry();

                Uri templateFileUri = new Uri(templateUrl);
                var webUrl = Web.WebUrlFromFolderUrlDirect(clientContext, templateFileUri);
                var templateContext = clientContext.Clone(webUrl.ToString());

                var library = templateUrl.ToLower().Replace(templateContext.Url.ToLower(), "").TrimStart('/');
                var idx = library.IndexOf("/", StringComparison.Ordinal);
                library = library.Substring(0, idx);

                // This syntax creates a SharePoint connector regardless we have the -InputInstance argument or not
                var fileConnector = new SharePointConnector(templateContext, templateContext.Url, library);
                string templateFileName = Path.GetFileName(templateUrl);
                XMLTemplateProvider provider = new XMLOpenXMLTemplateProvider(new OpenXMLConnector(templateFileName, fileConnector));
                templateFileName = templateFileName.Substring(0, templateFileName.LastIndexOf(".", StringComparison.Ordinal)) + ".xml";
                var provisioningTemplate = provider.GetTemplate(templateFileName, new ITemplateProviderExtension[0]);

                if (request.Parameters != null)
                {
                    foreach (var parameter in request.Parameters)
                    {
                        provisioningTemplate.Parameters[parameter.Key] = parameter.Value;
                    }
                }

                provisioningTemplate.Connector = provider.Connector;

                TokenReplaceCustomAction(provisioningTemplate, clientContext.Web);

                ProvisioningTemplateApplyingInformation applyingInformation = new ProvisioningTemplateApplyingInformation()
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        log.Info(String.Format("{0:00}/{1:00} - {2}", progress, total, message));
                    },
                    MessagesDelegate = (message, messageType) =>
                    {
                        log.Info(String.Format("{0} - {1}", messageType, message));
                    }
                };

                clientContext.Web.ApplyProvisioningTemplate(provisioningTemplate, applyingInformation);
                stopWatch.Stop();

                var applyTemplateResponse = new ApplyTemplateResponse 
                { 
                    TemplateApplied = true,
                    ElapsedMilliseconds = stopWatch.ElapsedMilliseconds
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<ApplyTemplateResponse>(applyTemplateResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                stopWatch.Stop();
                log.Error($"Error: {e.Message}\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        private static void TokenReplaceCustomAction(ProvisioningTemplate provisioningTemplate, Web web)
        {
            // List patch until PnP is updated
            if (provisioningTemplate.ClientSidePages == null) return;
            var tokenParser = new TokenParser(web, provisioningTemplate);
            foreach (var action in provisioningTemplate.CustomActions.SiteCustomActions)
            {
                if (action.ClientSideComponentProperties != null)
                    action.ClientSideComponentProperties = tokenParser.ParseString(action.ClientSideComponentProperties);
            }
            foreach (var action in provisioningTemplate.CustomActions.WebCustomActions)
            {
                if (action.ClientSideComponentProperties != null)
                    action.ClientSideComponentProperties = tokenParser.ParseString(action.ClientSideComponentProperties);
            }
        }

        public static void RedirectAssembly()
        {
            var list = AppDomain.CurrentDomain.GetAssemblies().OrderByDescending(a => a.FullName).Select(a => a.FullName).ToList();
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var requestedAssembly = new AssemblyName(args.Name);
                foreach (string asmName in list)
                {
                    if (asmName.StartsWith(requestedAssembly.Name + ","))
                    {
                        return Assembly.Load(asmName);
                    }
                }
                return null;
            };
        }

        public class ApplyTemplateRequest
        {
            [Required]
            [Display(Description = "URL of site to apply template")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "SPO URL to the .pnp template")]
            public string TemplateURL { get; set; }

            [Display(Description = "Replacement tokens to be used in .pnp templates")]
            public Parameter[] Parameters { get; set; }
        }

        public class Parameter
        {
            [Required]
            [Display(Description = "Extra PnP token to parse for template. Example. 'Foo' becomes '{parameter:Foo}' in the template.")]
            public string Key { get; set; }

            [Required]
            [Display(Description = "Value to replace with")]
            public string Value { get; set; }
        }


        public class ApplyTemplateResponse
        {
            [Display(Description = "True if template was applied")]
            public bool TemplateApplied { get; set; }

            [Display(Description = "Elapsed time in miliseconds")]
            public long ElapsedMilliseconds { get; set; }
        }
    }
}
