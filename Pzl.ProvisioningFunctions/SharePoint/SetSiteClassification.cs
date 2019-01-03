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
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SetSiteClassification
    {
        [FunctionName("SetSiteClassification")]
        [ResponseType(typeof(SetSiteClassificationResponse))]
        [Display(Name = "Set classification for the site", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteClassificationRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.SiteClassification))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteClassification");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                //clientContext.Site.SetSiteClassification(request.SiteClassification);
                clientContext.Site.Classification = request.SiteClassification;
                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteClassificationResponse>(new SetSiteClassificationResponse { ClassificationSet = true }, new JsonMediaTypeFormatter())
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

        public class SetSiteClassificationRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Site classification")]
            public string SiteClassification { get; set; }
        }

        public class SetSiteClassificationResponse
        {
            public bool ClassificationSet { get; set; }
        }
    }
}
