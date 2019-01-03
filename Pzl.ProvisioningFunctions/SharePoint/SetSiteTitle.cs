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
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class SetSiteTitle
    {
        [FunctionName("SetSiteTitle")]
        [ResponseType(typeof(SetSiteTitleResponse))]
        [Display(Name = "Set title for the site", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteTitleRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.Title))
                {
                    throw new ArgumentException("Parameter cannot be null", "Title");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                var web = clientContext.Web;
                web.Context.Load(web, w => w.Title);
                web.Context.ExecuteQueryRetry();

                var oldTitle = web.Title;
                if (oldTitle.Equals(request.Title))
                {
                    web.Title = request.Title;
                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteTitleResponse>(new SetSiteTitleResponse { OldTitle = oldTitle, NewTitle = request.Title }, new JsonMediaTypeFormatter())
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

        public class SetSiteTitleRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "")]
            public string Title { get; set; }
        }

        public class SetSiteTitleResponse
        {
            [Display(Description = "")]
            public string OldTitle { get; set; }
            [Display(Description = "")]
            public string NewTitle { get; set; }
        }
    }
}
