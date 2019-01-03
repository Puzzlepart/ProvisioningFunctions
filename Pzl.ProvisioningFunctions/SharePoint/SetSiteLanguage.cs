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
    public static class SetSiteLanguage
    {
        [FunctionName("SetSiteLanguage")]
        [ResponseType(typeof(SetSiteLanguageResponse))]
        [Display(Name = "Set allowed languages for the site", Description = "Which languages should be available by default for a site.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteLanguageRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (request.LanguageIds == null)
                {
                    throw new ArgumentException("Parameter cannot be null", "LanguageIds");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                var web = clientContext.Web;
                web.Context.Load(web, w => w.SupportedUILanguageIds);
                web.Update();
                web.Context.ExecuteQueryRetry();

                var isDirty = false;

                foreach (var lcid in web.SupportedUILanguageIds)
                {
                    var found = request.LanguageIds.Contains(lcid);

                    if (!found)
                    {
                        web.RemoveSupportedUILanguage(lcid);
                        isDirty = true;
                    }
                }
                if (isDirty)
                {
                    web.Update();
                    web.Context.ExecuteQueryRetry();
                }

                foreach (var lcid in request.LanguageIds)
                {
                    web.AddSupportedUILanguage(lcid);
                }
                web.Update();
                web.Context.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteLanguageResponse>(new SetSiteLanguageResponse { LanguagesModified = isDirty }, new JsonMediaTypeFormatter())
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

        public class SetSiteLanguageRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "LCID codes for languages. 1033 = English, 1044 = Norwegian")]
            public int[] LanguageIds { get; set; }
        }

        public class SetSiteLanguageResponse
        {
            [Display(Description = "True if languages was changed from the previous state")]
            public bool LanguagesModified { get; set; }
        }
    }
}
