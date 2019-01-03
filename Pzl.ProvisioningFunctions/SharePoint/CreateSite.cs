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
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Sites;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class CreateSite
    {
        [FunctionName("CreateSite")]
        [ResponseType(typeof(CreateSiteResponse))]
        [Display(Name = "Creates a modern team site", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]CreateSiteRequest request, TraceWriter log)
        {
            string adminUrl = $"https://{request.Tenant}-admin.sharepoint.com";

            try
            {

                if (string.IsNullOrWhiteSpace(request.Title))
                {
                    throw new ArgumentException("Parameter cannot be null", "Title");
                }
                if (string.IsNullOrWhiteSpace(request.Tenant))
                {
                    throw new ArgumentException("Parameter cannot be null", "Tenant");
                }
                if (string.IsNullOrWhiteSpace(request.Url))
                {
                    throw new ArgumentException("Parameter cannot be null", "Url");
                }

                var adminContext = await ConnectADAL.GetClientContext(adminUrl, log);
                Tenant tenant = new Tenant(adminContext);
                adminContext.ExecuteQuery();
                string url = $"https://{request.Tenant}.sharepoint.com/sites/{request.Url}";
                var siteCreationProperties = new SiteCreationProperties()
                {
                    Title = request.Title,
                    Url = url,
                    Owner = request.OwnerEmail,
                    Template = !string.IsNullOrWhiteSpace(request.Template) ? request.Template : "STS#3",
                    StorageMaximumLevel = 100,
                    UserCodeMaximumLevel = 0,
                    Lcid = request.Language != 0 ? request.Language : 1033,
                };
                tenant.CreateSite(siteCreationProperties);
                adminContext.ExecuteQuery();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<CreateSiteResponse>(new CreateSiteResponse { SiteURL = url }, new JsonMediaTypeFormatter())
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

        public class CreateSiteRequest
        {
            [Required]
            [Display(Description = "Tenant")]
            public string Tenant { get; set; }
            [Required]
            [Display(Description = "Title ")]
            public string Title { get; set; }
            [Required]
            [Display(Description = "Url")]
            public string Url { get; set; }
            [Display(Description = "Description")]
            public string Description { get; set; }
            [Display(Description = "OwnerEmail")]
            public string OwnerEmail { get; set; }
            [Display(Description = "Template")]
            public string Template { get; set; }
            [Display(Description = "Language ")]
            public uint Language { get; set; }
        }

        public class CreateSiteResponse
        {

            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }
    }
}
