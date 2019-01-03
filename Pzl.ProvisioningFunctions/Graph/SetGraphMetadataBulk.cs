using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.Graph
{
    public static class SetGraphMetadataBulk
    {
        [FunctionName("SetGraphMetadataBulk")]
        [ResponseType(typeof(SetGraphMetadataBulkResponse))]
        [Display(Name = "Set Office 365 Group generic metadata as a JSON payload", Description = "Store metadata using techmikael_GenericSchema for the Office 365 Group in the Microsoft Graph")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGraphMetadataBulkRequest request, TraceWriter log)
        {
            try
            {
                const string extensionName = "techmikael_GenericSchema";

                var payload = $"{{\"{extensionName}\": {request.JSON}}}";

                var content = new StringContent(payload, Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/v1.0/groups/{request.GroupId}");
                string bearerToken = await ConnectADAL.GetBearerToken();
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PatchAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetGraphMetadataBulkResponse>(new SetGraphMetadataBulkResponse { Added = true }, new JsonMediaTypeFormatter())
                    });
                }
                string responseMsg = await response.Content.ReadAsStringAsync();
                dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<object>(errorMsg, new JsonMediaTypeFormatter())
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


        public class SetGraphMetadataBulkRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = @"Escaped JSON string which match techmikael_GenericSchema in the format:

{
	""KeyString01"": ""Title"",
	""ValueString01"": ""CTO Puzzlepart"",

	""KeyBoolean00"": ""IsMVP"",
	""LabelBoolean00"": ""Is Microsoft MVP"",
	""ValueBoolean00"": true
}
")]
            public string JSON { get; set; }
        }

        public class SetGraphMetadataBulkResponse
        {
            [Display(Description = "true/false if added")]
            public bool Added { get; set; }
        }
    }
}
