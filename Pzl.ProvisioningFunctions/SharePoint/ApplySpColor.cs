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
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class ApplySpColor
    {
        [FunctionName("ApplySpColor")]
        [ResponseType(typeof(ApplySpColorResponse))]
        [Display(Name = "Apply .spcolor file to the site", Description = "Apply a previously added .spcolor file to the site.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]ApplySpColorRequest request, TraceWriter log)
        {
            try
            {
                string siteUrl = request.SiteURL;
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                var web = clientContext.Web;
                web.Lists.EnsureSiteAssetsLibrary();
                string relativeSiteUrl = UrlUtility.MakeRelativeUrl(siteUrl);
                string siteAssetsUrl = UrlUtility.Combine(relativeSiteUrl, "SiteAssets");
                var fileUrl = UrlUtility.Combine(siteAssetsUrl, request.Filename);

                web.ApplyTheme(fileUrl, null, null, true);
                web.Update();
                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<ApplySpColorResponse>(new ApplySpColorResponse { Applied = true }, new JsonMediaTypeFormatter())
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

        public class ApplySpColorRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Filename of spcolor file in SiteAssets")]
            public string Filename { get; set; }
        }

        public class ApplySpColorResponse
        {
            [Display(Description = "True if applied")]
            public bool Applied { get; set; }
        }
    }
}
