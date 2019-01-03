using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Web.Http.Description;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Pzl.ProvisioningFunctions.Helpers;
using Group = Microsoft.Graph.Group;
using System.Linq;
namespace Pzl.ProvisioningFunctions.Graph
{
    public static class DeleteO365GroupAndSPSite
    {
        [FunctionName("DeleteO365GroupAndSPSite")]

        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]DeleteGroupRequest request, TraceWriter log)
        {
            try
            {
              
                GraphServiceClient client = ConnectADAL.GetGraphClient(GraphEndpoint.Beta);

                // get the group to be deleted
                var delgroup = await client.Groups[request.GroupId].Request().GetAsync();

                // Create a response for the deleted group
                var DelGroupResponse = new DeletedGroupResponse
                {
                    GroupId = delgroup.Id,
                    DisplayName = delgroup.DisplayName,
                    Mail = delgroup.Mail

                };

                // delete the group, if the group is deleted the SP Site will be deleted also but take some time.
                await client.Groups[request.GroupId].Request().DeleteAsync();


                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<DeletedGroupResponse>(DelGroupResponse, new JsonMediaTypeFormatter())
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

        public class DeleteGroupRequest
        {
            [Required]
            [Display(Description = "GroupId to be deleted")]
            public string GroupId { get; set; }            
        }

        public class DeletedGroupResponse
        {
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Display(Description = "DisplayName of the Office 365 Group")]
            public string DisplayName { get; set; }

             [Display(Description = "Mail of the Office 365 Group")]
            public string Mail { get; set; }

        }





    }
}
