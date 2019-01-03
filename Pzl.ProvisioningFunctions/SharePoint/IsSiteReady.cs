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
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class IsSiteReady
    {
        [FunctionName("IsSiteReady")]
        [ResponseType(typeof(IsSiteReadyResponse))]
        [Display(Name = "Check if Modern Team Site is ready", Description = "Check if the team site is 100% ready before running operations against it.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]IsSiteReadyRequest request, TraceWriter log)
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
                var clientContext = await ConnectADAL.GetClientContext(adminUrl, log);

                Tenant tenant = new Tenant(clientContext);
                var site = tenant.GetSitePropertiesByUrl(siteUrl, false);
                clientContext.Load(site, s => s.Status);
                site.Context.ExecuteQueryRetry();
                var status = site.Status;
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<IsSiteReadyResponse>(new IsSiteReadyResponse{IsSiteReady = status == "Active"}, new JsonMediaTypeFormatter())
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

        public class IsSiteReadyRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }

        public class IsSiteReadyResponse
        {
            [Display(Description = "True/false if site is ready to accept modifications via API's")]
            public bool IsSiteReady { get; set; }
        }
    }
}
