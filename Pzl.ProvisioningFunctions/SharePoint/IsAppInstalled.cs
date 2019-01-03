using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class IsAppInstalled
    {
        [FunctionName("IsAppInstalled")]
        [ResponseType(typeof(IsAppInstalledResponse))]
        [Display(Name = "Check if an SP App is installed", Description = "Checks if an application is completely installed.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]IsAppInstalledRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;
            try
            {
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                var web = clientContext.Web;

                bool isInstalled = false;
                var instances = web.GetAppInstances();
                if (instances != null)
                {
                    isInstalled = instances.Any(i => i.Status == AppInstanceStatus.Installed
                                                     && i.Title.Equals(request.Title, StringComparison.InvariantCultureIgnoreCase));
                }

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<IsAppInstalledResponse>(new IsAppInstalledResponse { Installed = isInstalled }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class IsAppInstalledRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Title of application")]
            public string Title { get; set; }
        }

        public class IsAppInstalledResponse
        {
            [Display(Description = "True if app is installed")]
            public bool Installed { get; set; }
        }
    }
}
