using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SetNoScript
    {
        [FunctionName("SetNoScript")]
        [ResponseType(typeof(SetNoScriptResponse))]
        [Display(Name = "Enable or disable NoScript", Description = "Turn NoScript on or off for the site collection")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetNoScriptRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;
            Uri uri = new Uri(siteUrl);
            string adminUrl = uri.Scheme + "://" + uri.Authority;
            if (!adminUrl.Contains("-admin.sharepoint.com"))
            {
                adminUrl = adminUrl.Replace(".sharepoint.com", "-admin.sharepoint.com");
            }

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }

                var clientContext = await ConnectADAL.GetClientContext(adminUrl, log);

                Tenant tenant = new Tenant(clientContext);
                var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, false);
                clientContext.Load(siteProperties, s => s.SharingCapability);
                siteProperties.Context.ExecuteQueryRetry();
                siteProperties.DenyAddAndCustomizePages = (request.NoScript ? DenyAddAndCustomizePagesStatus.Enabled : DenyAddAndCustomizePagesStatus.Disabled);
                siteProperties.Update();
                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetNoScriptResponse>(new SetNoScriptResponse { UpdatedNoScript = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class SetNoScriptRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Turn NoScript on or off")]
            public bool NoScript { get; set; }
        }

        public class SetNoScriptResponse
        {
            [Display(Description = "True/false if NoScript updated")]
            public bool UpdatedNoScript { get; set; }
        }
    }
}
