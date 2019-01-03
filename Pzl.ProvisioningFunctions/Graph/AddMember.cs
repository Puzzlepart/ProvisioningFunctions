using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.Graph
{
    public static class AddMember
    {
        [FunctionName("AddMember")]
        [Display(Name = "Add a member to an Office 365 Group", Description = "This action will add an member to an Office 365 Group")]
        public static async Task<AddMemberResponse> Run([HttpTrigger(AuthorizationLevel.Function, "post")]AddMemberRequest request, TraceWriter log)
        {
            GraphServiceClient client = ConnectADAL.GetGraphClient();

            int idx = request.LoginName.LastIndexOf('|');
            if (idx > 0)
            {
                request.LoginName = request.LoginName.Substring(idx + 1);
            }

            var memberQuery = await client.Users
                .Request()
                .Filter($"userPrincipalName eq '{request.LoginName}'")
                .GetAsync();

            var member = memberQuery.FirstOrDefault();
            bool added = false;
            if (member != null)
            {
                try
                {
                    // And if any, add it to the collection of group's members
                    await client.Groups[request.GroupId].Members.References.Request().AddAsync(member);
                    added = true;
                }
                catch (ServiceException ex)
                {
                    if (ex.Error.Code == "Request_BadRequest" &&
                        ex.Error.Message.Contains("added object references already exist"))
                    {
                        // Skip any already existing member
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            return new AddMemberResponse() { Added = added };
        }

        public class AddMemberRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Unique login name of the member (user principal name)")]
            public string LoginName { get; set; }
        }

        public class AddMemberResponse
        {
            [Display(Description = "true/false if added")]
            public bool Added { get; set; }
        }
    }
}
