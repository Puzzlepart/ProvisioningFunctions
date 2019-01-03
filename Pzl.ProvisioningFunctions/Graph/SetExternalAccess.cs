using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.Graph
{
    public static class AllowExternalMembersInGroup
    {
        [FunctionName("AllowExternalMembersInGroup")]
        [Display(Name = "Enable or disable invitation of external members", Description = "Allow or disallow invitation of external members to the Office 365 Group")]
        public static async Task<AllowExternalMembersInGroupResponse> Run([HttpTrigger(AuthorizationLevel.Function, "post")]AllowExternalMembersInGroupRequest request, TraceWriter log)
        {
            GraphServiceClient client = ConnectADAL.GetGraphClient();
            try
            {
                const string externalTemplateId = "08d542b9-071f-4e16-94b0-74abb372e3d9";
                GroupSetting externalMemberSetting = new GroupSetting { TemplateId = externalTemplateId };
                SettingValue setVal = new SettingValue
                {
                    Name = "AllowToAddGuests",
                    Value = request.ExternalAllowed.ToString()
                };
                externalMemberSetting.Values = new List<SettingValue> { setVal };

                var existingSettings = await client.Groups[request.GroupId].Settings.Request().GetAsync();

                bool hasExistingSetting = false;
                foreach (GroupSetting groupSetting in existingSettings)
                {
                    if (!groupSetting.TemplateId.Equals(externalTemplateId, StringComparison.InvariantCultureIgnoreCase)) continue;

                    await client.Groups[request.GroupId].Settings[groupSetting.Id].Request().UpdateAsync(externalMemberSetting);
                    hasExistingSetting = true;
                    break;
                }

                if (!hasExistingSetting)
                {
                    await client.Groups[request.GroupId].Settings.Request().AddAsync(externalMemberSetting);
                }

                return new AllowExternalMembersInGroupResponse { ExternalAllowed = request.ExternalAllowed };
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                throw;
            }
        }

        public class AllowExternalMembersInGroupRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Display(Description = "Specify if external users can be added as members")]
            public bool ExternalAllowed { get; set; }
        }

        public class AllowExternalMembersInGroupResponse
        {
            [Display(Description = "True/false if external users are allowed as members")]
            public bool ExternalAllowed { get; set; }
        }
    }
}
