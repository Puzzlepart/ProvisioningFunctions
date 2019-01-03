using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
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
    public static class SetGraphMetadataGeneric
    {
        [FunctionName("SetGraphMetadataGeneric")]
        [ResponseType(typeof(SetGraphMetadataGenericResponse))]
        [Display(Name = "Set Office 365 Group generic metadata", Description = "Store metadata using techmikael_GenericSchema for the Office 365 Group in the Microsoft Graph")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGraphMetadataGenericRequest request, TraceWriter log)
        {
            try
            {
                const string extensionName = "techmikael_GenericSchema";

                string propertyName = request.Name;

                dynamic property = new ExpandoObject();
                ((IDictionary<string, object>)property).Add(propertyName, request.Value);

                dynamic template = new ExpandoObject();
                ((IDictionary<string, object>)template).Add(extensionName, property);

                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/v1.0/groups/{request.GroupId}");
                string bearerToken = await ConnectADAL.GetBearerToken();
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PatchAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetGraphMetadataGenericResponse>(new SetGraphMetadataGenericResponse { Added = true }, new JsonMediaTypeFormatter())
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


        public class SetGraphMetadataGenericRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Metadata name. E.g.: ValueString00, KeyString00, LabelString00")]
            public string Name { get; set; }

            [Required]
            [Display(Description = "Metadata value. The actual value to be stored in e.g. ValueString00")]
            public string Value { get; set; }
        }

        public class SetGraphMetadataGenericResponse
        {
            [Display(Description = "true/false if added")]
            public bool Added { get; set; }
        }
    }
}
