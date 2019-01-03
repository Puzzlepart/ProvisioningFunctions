using System;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.Graph
{
    public static class GetGroupSite
    {
        [FunctionName("GetGroupSite")]
        [Display(Name = "Get site URL for Office 365 Group", Description = "Retreive the URL to the Modern Team Site associated with the Office 365 Group")]
        public static async Task<GetGroupSiteResponse> Run([HttpTrigger(AuthorizationLevel.Function, "post")]GetGroupSiteRequest request, TraceWriter log)
        {
            GraphServiceClient client = ConnectADAL.GetGraphClient();
            try
            {
                var rootSite = await client.Groups[request.GroupId].Sites["root"].Request().GetAsync();
                return string.IsNullOrWhiteSpace(rootSite.WebUrl) ? new GetGroupSiteResponse { SiteURL = "n/a" } : new GetGroupSiteResponse { SiteURL = rootSite.WebUrl };
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return new GetGroupSiteResponse(){SiteURL = e.Message};
            }
        }

        public class GetGroupSiteRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class GetGroupSiteResponse
        {
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }
    }
}
