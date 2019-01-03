using System;
using System.ComponentModel.DataAnnotations;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.Graph
{
    public static class GetClassificationValues
    {
        [FunctionName("GetClassificationValues")]
        [Display(Name = "Retrieve classification values", Description = "Fetch the classification values defined on the tenant for Office 365 Groups")]
        public static async Task<GetClassificationValuesResponse> Run([HttpTrigger(AuthorizationLevel.Function, "get")]HttpRequestMessage req, TraceWriter log)
        {
            GraphServiceClient client = ConnectADAL.GetGraphClient();
            try
            {
                const string groupTemplateId = "62375ab9-6b52-47ed-826b-58e47e0e304b";
                var existingSettings = await client.GroupSettings.Request().GetAsync();

                var response = new GetClassificationValuesResponse();
                foreach (GroupSetting groupSetting in existingSettings)
                {
                    if (!groupSetting.TemplateId.Equals(groupTemplateId, StringComparison.InvariantCultureIgnoreCase)) continue;
                    foreach (SettingValue settingValue in groupSetting.Values)
                    {
                        if (settingValue.Name.Equals("ClassificationList"))
                        {
                            response.Classifications = settingValue.Value.Split(',');
                        }
                        else if (settingValue.Name.Equals("DefaultClassification"))
                        {
                            response.DefaultClassification = settingValue.Value;
                        }
                    }
                    break;
                }

                return response;
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                throw;
            }
        }

        public class GetClassificationValuesResponse
        {
            [Display(Description = "Default classification value if set")]
            public string DefaultClassification { get; set; }

            [Display(Description = "List of classification values for the tenant")]
            public string[] Classifications { get; set; }

        }
    }
}
