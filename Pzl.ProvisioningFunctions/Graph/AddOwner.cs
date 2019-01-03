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
    public static class AddOwner
    {
        [FunctionName("AddOwner")]
        [Display(Name = "Add an owner to an Office 365 Group", Description = "This action will add an owner to an Office 365 Group")]
        public static async Task<AddOwnerResponse> Run([HttpTrigger(AuthorizationLevel.Function, "post")]AddOwnerRequest request, TraceWriter log)
        {
            GraphServiceClient client = ConnectADAL.GetGraphClient();

            int idx = request.LoginName.LastIndexOf('|');
            if (idx > 0)
            {
                request.LoginName = request.LoginName.Substring(idx+1);
            }

            var ownerQuery = await client.Users
                .Request()
                .Filter($"userPrincipalName eq '{request.LoginName}'")
                .GetAsync();

            var owner = ownerQuery.FirstOrDefault();
            bool added = false;
            if (owner != null)
            {
                try
                {
                    // And if any, add it to the collection of group's owners
                    log.Info($"Adding user {request.LoginName} to Owners group for {request.GroupId}");
                    await client.Groups[request.GroupId].Owners.References.Request().AddAsync(owner);
                    if (request.AddOption == AddOwnerOption.AddAsOwnerAndMember)
                    {
                        log.Info($"Adding user {request.LoginName} to Members group for {request.GroupId}");
                        await client.Groups[request.GroupId].Members.References.Request().AddAsync(owner);
                    }
                    added = true;
                }
                catch (ServiceException ex)
                {
                    if (ex.Error.Code == "Request_BadRequest" &&
                        ex.Error.Message.Contains("added object references already exist"))
                    {
                        // Skip any already existing owner
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            return new AddOwnerResponse() { Added = added };
        }

        public enum AddOwnerOption
        {
            [Display(Name = "Add as Owner only")]
            AddAsOwnerOnly,
            [Display(Name = "Add as Owner and Member")]
            AddAsOwnerAndMember,
        }

        public class AddOwnerRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Unique login name of the owner (user principal name)")]
            public string LoginName { get; set; }
            [Display(Description = "Add option")]
            public AddOwnerOption AddOption { get; set; }
        }

        public class AddOwnerResponse
        {
            [Display(Description = "true/false if added")]
            public bool Added { get; set; }
        }
    }
}
