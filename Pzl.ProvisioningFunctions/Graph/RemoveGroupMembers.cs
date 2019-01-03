using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Pzl.ProvisioningFunctions.Helpers;
using Microsoft.Graph;

namespace Pzl.ProvisioningFunctions.Graph
{
    public class GroupUser
    {
        public string Id { get; set; }
        public string UserType { get; set; }
        public string Mail { get; set; }
    }

    public static class RemoveGroupMembers
    {
        [FunctionName("RemoveGroupMembers")]
        [ResponseType(typeof(RemoveGroupMembersResponse))]
        [Display(Name = "Remove group members", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]RemoveGroupMembersRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.GroupId))
                {
                    throw new ArgumentException("Parameter cannot be null", "GroupId");
                }
                GraphServiceClient client = ConnectADAL.GetGraphClient(GraphEndpoint.v1);
                var group = client.Groups[request.GroupId];
                var members = await group.Members.Request().Select("displayName, id, mail, userPrincipalName, userType").GetAsync();

                var users = new List<GroupUser>();

                foreach (var u in members.CurrentPage.Where(p => p.GetType() == typeof(User)).Cast<User>().ToList())
                {
                    users.Add(new GroupUser() { Id = u.Id, UserType = u.UserType, Mail = u.Mail });
                }


                // Removing users from group members
                for (int i = 0; i < users.Count; i++)
                {
                    var user = users[i];
                    if(user.UserType != "Guest" && request.GroupMembersRemoval == GroupMembersRemoval.GuestUsersOnly)
                    {
                        continue;
                    }                   
                    log.Info($"Removing user {user.Id} from group {request.GroupId}");
                    await group.Members[user.Id].Reference.Request().DeleteAsync();
                }


                var removedGuestUsers = new List<GroupUser>();
                if (request.RemoveGuestUsers)
                {
                    var guestUsers = users.Where(u => u.UserType.Equals("Guest")).ToList();
                    for (int i = 0; i < guestUsers.Count; i++)
                    {
                        var guestUser = guestUsers[i];
                        log.Info($"Retrieving unified membership for user {guestUser.Id}");
                        var memberOfPage = await client.Users[guestUser.Id].MemberOf.Request().GetAsync();
                        var unifiedGroups = memberOfPage.CurrentPage.Where(p => p.GetType() == typeof(Group)).Cast<Group>().ToList().Where(g => g.GroupTypes.Contains("Unified")).ToList();
                        if (unifiedGroups.Count == 0)
                        {
                            log.Info($"Removing guest user {guestUser.Id}");
                            await client.Users[guestUser.Id].Request().DeleteAsync();
                            removedGuestUsers.Add(guestUser);
                        }
                    }
                }

                var removeGroupMembersResponse = new RemoveGroupMembersResponse {
                    RemovedMembers = users,
                    RemovedGuestUsers = removedGuestUsers
                };

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<RemoveGroupMembersResponse>(removeGroupMembersResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public static async Task RemoveUsersFromGroup(string groupId, IGroupRequestBuilder group, List<User> users, TraceWriter log)
        {
            for (int i = 0; i < users.Count; i++)
            {
                var user = users[i];
                log.Info($"Removing user {user.Id} from group {groupId}");
                await group.Members[user.Id].Reference.Request().DeleteAsync();
            }
        }

        public class RemoveGroupMembersRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
            [Display(Description = "Should guest users with no remaining Unified membership be removed from AD")]
            public bool RemoveGuestUsers { get; set; }
            [Display(Description = "Group members removal")]
            public GroupMembersRemoval GroupMembersRemoval { get; set; }
        }

        public class RemoveGroupMembersResponse
        {
            [Display(Description = "List of removed members")]
            public List<GroupUser> RemovedMembers { get; set; }
            [Display(Description = "List of removed guest users")]
            public List<GroupUser> RemovedGuestUsers { get; set; }
        }

        public enum GroupMembersRemoval
        {
            [Display(Name = "Remove all group members")]
            All,
            [Display(Name = "Remove all guest group members")]
            GuestUsersOnly,
        }
    }
}
