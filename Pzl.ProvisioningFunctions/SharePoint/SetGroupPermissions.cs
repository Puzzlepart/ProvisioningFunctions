using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SetGroupPermissions
    {
        [FunctionName("SetGroupPermissions")]
        [ResponseType(typeof(SetGroupPermissionsResponse))]
        [Display(Name = "Set group permissions", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGroupPermissionsRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);


                var web = clientContext.Web;
                var associatedOwnerGroup = web.AssociatedOwnerGroup;
                var associatedMemberGroup = web.AssociatedMemberGroup;
                var associatedVisitorGroup = web.AssociatedVisitorGroup;
                var webRoleDefinitions = web.RoleDefinitions;
                var webRoleAssignments = web.RoleAssignments;
                web.Context.Load(associatedOwnerGroup);
                web.Context.Load(associatedMemberGroup);
                web.Context.Load(associatedVisitorGroup);
                web.Context.Load(webRoleDefinitions);
                web.Context.Load(webRoleAssignments);
                web.Context.ExecuteQueryRetry();

                var associatedOwnerGroupRoleAss = webRoleAssignments.Where(roleAss => roleAss.PrincipalId == associatedOwnerGroup.Id).ToList();
                var associatedMemberGroupRoleAss = webRoleAssignments.Where(roleAss => roleAss.PrincipalId == associatedMemberGroup.Id).ToList();
                var associatedVisitorGroupRoleAss = webRoleAssignments.Where(roleAss => roleAss.PrincipalId == associatedVisitorGroup.Id).ToList();

                for (var i = 0; i < associatedOwnerGroupRoleAss.Count; i++)
                {
                    associatedOwnerGroupRoleAss[i].DeleteObject();
                }

                for (var i = 0; i < associatedMemberGroupRoleAss.Count; i++)
                {
                    associatedMemberGroupRoleAss[i].DeleteObject();
                }

                for (var i = 0; i < associatedVisitorGroupRoleAss.Count; i++)
                {
                    associatedVisitorGroupRoleAss[i].DeleteObject();
                }

                web.Update();
                web.Context.ExecuteQueryRetry();

                var associatedOwnerGroupRdb = new RoleDefinitionBindingCollection(clientContext)
                {
                    webRoleDefinitions.GetByType(request.OwnersPermissionLevel)
                };
                webRoleAssignments.Add(associatedOwnerGroup, associatedOwnerGroupRdb);

                var associatedMemberGroupRdp = new RoleDefinitionBindingCollection(clientContext)
                {
                    webRoleDefinitions.GetByType(request.MembersPermissionLevel)
                };
                webRoleAssignments.Add(associatedMemberGroup, associatedMemberGroupRdp);

                var associatedVisitorGroupRdb = new RoleDefinitionBindingCollection(clientContext)
                {
                    webRoleDefinitions.GetByType(request.VisitorsPermissionLevel)
                };
                webRoleAssignments.Add(associatedVisitorGroup, associatedVisitorGroupRdb);

                web.Update();
                web.Context.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetGroupPermissionsResponse>(new SetGroupPermissionsResponse { PermissionsModified = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetGroupPermissionsResponse>(new SetGroupPermissionsResponse { PermissionsModified = false }, new JsonMediaTypeFormatter())
                });
            }
        }

        public class SetGroupPermissionsRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
            [Required]
            [Display(Description = "Permission level for Owners")]
            public RoleType OwnersPermissionLevel { get; set; }
            [Required]
            [Display(Description = "Permission level for Members")]
            public RoleType MembersPermissionLevel { get; set; }
            [Required]
            [Display(Description = "Permission level for Visitors")]
            public RoleType VisitorsPermissionLevel { get; set; }

        }

        public class SetGroupPermissionsResponse
        {

            [Display(Description = "Was group permissions modified")]
            public bool PermissionsModified { get; set; }
        }
    }
}
