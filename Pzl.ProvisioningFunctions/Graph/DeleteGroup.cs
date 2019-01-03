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
using Cumulus.Monads.Helpers;
using Group = Microsoft.Graph.Group;

namespace Cumulus.Monads.Graph
{
    public static class DeleteGroup
    {
        [FunctionName("DeleteGroup")]
        [ResponseType(typeof(DeleteGroupResponse))]
        [Display(Name = "Delete Office 365 Group", Description = "This action will delete a Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]DeleteGroupRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.GroupId))
                {
                    throw new ArgumentException("Parameter cannot be null", "GroupId");
                }

                GraphServiceClient client = ConnectADAL.GetGraphClient(GraphEndpoint.v1);
                await client.Groups[request.GroupId].Request().DeleteAsync();
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<DeleteGroupResponse>(new DeleteGroupResponse { GroupDeleted = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<DeleteGroupResponse>(new DeleteGroupResponse { GroupDeleted = false }, new JsonMediaTypeFormatter())
                });
            }
        }

        public class DeleteGroupRequest
        {
            [Required]
            [Display(Description = "Group ID")]
            public string GroupId { get; set; }
        }

        public class DeleteGroupResponse
        {
            [Display(Description = "True if the group was deleted")]
            public bool GroupDeleted { get; set; }
        }
    }
}
