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
using Microsoft.SharePoint.Client.Taxonomy;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class GetTermProperty
    {
        [FunctionName("GetTermProperty")]
        [ResponseType(typeof(GetTermPropertyResponse))]
        [Display(Name = "Get term property", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]GetTermPropertyRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.TermGUID))
                {
                    throw new ArgumentException("Parameter cannot be null", "TermGUID");
                }
                if (string.IsNullOrWhiteSpace(request.PropertyName))
                {
                    throw new ArgumentException("Parameter cannot be null", "PropertyName");
                }
                if (string.IsNullOrWhiteSpace(request.FallbackValue))
                {
                    request.FallbackValue = "";
                }
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                clientContext.Load(taxonomySession);
                clientContext.ExecuteQueryRetry();
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                var term = termStore.GetTerm(new Guid(request.TermGUID));
                clientContext.Load(term, t => t.LocalCustomProperties);
                clientContext.ExecuteQueryRetry();
                var propertyValue = string.Empty;
                do
                {
                    if (term.LocalCustomProperties.Keys.Contains(request.PropertyName))
                    {
                        propertyValue = term.LocalCustomProperties[request.PropertyName];
                    }
                    else
                    {
                        term = term.Parent;
                        clientContext.Load(term, t => t.LocalCustomProperties);
                        clientContext.ExecuteQueryRetry();
                    }

                } while (string.IsNullOrWhiteSpace(propertyValue));

                var getTermPropertyResponse = new GetTermPropertyResponse
                {
                    PropertyValue = propertyValue
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<GetTermPropertyResponse>(getTermPropertyResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception)
            {
                var getTermPropertyResponse = new GetTermPropertyResponse
                {
                    PropertyValue = request.FallbackValue
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<GetTermPropertyResponse>(getTermPropertyResponse, new JsonMediaTypeFormatter())
                });
            }
        }

        public class GetTermPropertyRequest
        {
            [Required]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Term GUID")]
            public string TermGUID { get; set; }

            [Required]
            [Display(Description = "Property name")]
            public string PropertyName { get; set; }

            [Display(Description = "Fallback value")]
            public string FallbackValue { get; set; }
        }

        public class GetTermPropertyResponse
        {
            public string PropertyValue { get; set; }
        }
    }
}
