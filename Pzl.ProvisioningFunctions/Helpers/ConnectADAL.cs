using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using HttpUtility = System.Web.HttpUtility;

namespace Pzl.ProvisioningFunctions.Helpers
{
    enum GraphEndpoint
    {
        v1,
        Beta
    }
    class ConnectADAL
    {
        private static readonly Uri ADALLogin = new Uri("https://login.windows.net/");
        const string GraphResourceId = "https://graph.microsoft.com"; // Microsoft Graph End-point
        private static string _appId = Environment.GetEnvironmentVariable("ADALAppId");
        private static string _appSecret = Environment.GetEnvironmentVariable("ADALAppSecret");
        private static string _appCert = Environment.GetEnvironmentVariable("ADALAppCertificate");
        private static string _appCertKey = Environment.GetEnvironmentVariable("ADALAppCertificateKey");
        private static string _tenantAdmin = Environment.GetEnvironmentVariable("TenantAdmin");
        private static string _tenantPassword = Environment.GetEnvironmentVariable("TenantPassword");

        private static readonly string ADALDomain = Environment.GetEnvironmentVariable("ADALDomain");
        private static readonly Dictionary<string, AuthenticationResult> ResourceTokenLookup = new Dictionary<string, AuthenticationResult>();

        private static readonly AzureServiceTokenProvider AzureServiceTokenProvider = new AzureServiceTokenProvider();
        private static readonly KeyVaultClient KvClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(AzureServiceTokenProvider.KeyVaultTokenCallback));

        public class MsiInformation
        {
            public string OwnerId { get; set; }
            public string BearerToken { get; set; }
        }

        private static string SecretUri(string secret)
        {
            return $"{ConfigurationManager.AppSettings["KeyVaultUri"].TrimEnd('/')}/Secrets/{secret}";
        }

        private static async Task GetVariables()
        {
            //http://integration.team/2017/09/25/retrieve-azure-key-vault-secrets-using-azure-functions-managed-service-identity/
            if (string.IsNullOrEmpty(_appId)) _appId = (await KvClient.GetSecretAsync(SecretUri("ADALAppId"))).Value;
            if (string.IsNullOrEmpty(_appSecret)) _appSecret = (await KvClient.GetSecretAsync(SecretUri("ADALAppSecret"))).Value;
            if (string.IsNullOrEmpty(_appCert)) _appCert = (await KvClient.GetSecretAsync(SecretUri("ADALAppCertificate"))).Value;
            if (string.IsNullOrEmpty(_appCertKey)) _appCertKey = (await KvClient.GetSecretAsync(SecretUri("ADALAppCertificateKey"))).Value;

            try
            {
                if (string.IsNullOrEmpty(_tenantAdmin)) _tenantAdmin = (await KvClient.GetSecretAsync(SecretUri("TenantAdmin"))).Value;
                if (string.IsNullOrEmpty(_tenantPassword)) _tenantPassword = (await KvClient.GetSecretAsync(SecretUri("TenantPassword"))).Value;
            }
            catch (Exception)
            {
                //silent catch if no admin/pw as this is needed only for lifecycle policy atm
            }
        }

        private static async Task<string> GetAccessToken(string AADDomain, bool useTenantAdmin = false)
        {
            await GetVariables();
            AuthenticationResult token = null;
            if (!useTenantAdmin && ResourceTokenLookup.TryGetValue(GraphResourceId, out token) &&
                token.ExpiresOn.UtcDateTime >= DateTime.UtcNow.AddMinutes(-5))
            {
                //Return cached token for ADAL app-only tokens
                return token.AccessToken;
            }

            var authenticationContext = new AuthenticationContext(ADALLogin + AADDomain);
            bool keepRetry = false;
            do
            {
                TimeSpan? delay = null;
                try
                {
                    if (!useTenantAdmin)
                    {
                        var clientCredential = new ClientCredential(_appId, _appSecret);
                        token = await authenticationContext.AcquireTokenAsync(GraphResourceId, clientCredential);
                    }
                    else
                    {
                        // Hack to get user token from a Web/API adal app - passing both username/password and client secret
                        // Ref AADSTS70002 error
                        Uri authUri = new Uri($"{ADALLogin}{ADALDomain}/oauth2/token");
                        var contentString = $"resource={HttpUtility.UrlEncode(GraphResourceId)}&client_id={_appId}&client_secret={_appSecret}&grant_type=password&username={HttpUtility.UrlEncode(_tenantAdmin)}&password={HttpUtility.UrlEncode(_tenantPassword)}&scope=openid";
                        var content = new StringContent(contentString, Encoding.UTF8, "application/x-www-form-urlencoded");
                        HttpClient client = new HttpClient();
                        var response = await client.PostAsync(authUri, content);

                        string responseMsg = await response.Content.ReadAsStringAsync();
                        if (response.IsSuccessStatusCode)
                        {
                            JObject tokendata = JsonConvert.DeserializeObject<JObject>(responseMsg);
                            return tokendata["access_token"].ToString();
                        }
                        throw new Exception(responseMsg);
                    }
                }
                catch (Exception ex)
                {
                    if (!(ex is AdalServiceException) && !(ex.InnerException is AdalServiceException)) throw;

                    AdalServiceException serviceException;
                    if (ex is AdalServiceException) serviceException = (AdalServiceException)ex;
                    else serviceException = (AdalServiceException)ex.InnerException;
                    if (serviceException.ErrorCode == "temporarily_unavailable")
                    {
                        RetryConditionHeaderValue retry = serviceException.Headers.RetryAfter;
                        if (retry.Delta.HasValue)
                        {
                            delay = retry.Delta;
                        }
                        else if (retry.Date.HasValue)
                        {
                            delay = retry.Date.Value.Offset;
                        }
                        if (delay.HasValue)
                        {
                            Thread.Sleep((int)delay.Value.TotalSeconds); // sleep or other
                            keepRetry = true;
                        }
                    }
                    else
                    {
                        throw;
                    }
                }
            } while (keepRetry);

            ResourceTokenLookup[GraphResourceId] = token;
            return token.AccessToken;
        }

        private static async Task<string> GetAccessTokenSharePoint(string AADDomain, string siteUrl, TraceWriter log = null)
        {
            await GetVariables();
            //https://blogs.msdn.microsoft.com/richard_dizeregas_blog/2015/05/03/performing-app-only-operations-on-sharepoint-online-through-azure-ad/
            AuthenticationResult token;
            Uri uri = new Uri(siteUrl);
            string resourceUri = uri.Scheme + "://" + uri.Authority;
            if (ResourceTokenLookup.TryGetValue(resourceUri, out token) &&
                token.ExpiresOn.UtcDateTime >= DateTime.UtcNow.AddMinutes(-5))
            {
                return token.AccessToken;
            }
            if (token != null)
            {
                log?.Info($"Token expired {token.ExpiresOn.UtcDateTime}");
            }

            var cac = GetClientAssertionCertificate();
            var authenticationContext = new AuthenticationContext(ADALLogin + AADDomain);

            bool keepRetry = false;
            do
            {
                TimeSpan? delay = null;
                try
                {
                    token = await authenticationContext.AcquireTokenAsync(resourceUri, cac);
                }
                catch (Exception ex)
                {
                    if (!(ex is AdalServiceException) && !(ex.InnerException is AdalServiceException)) throw;

                    AdalServiceException serviceException;
                    if (ex is AdalServiceException) serviceException = (AdalServiceException)ex;
                    else serviceException = (AdalServiceException)ex.InnerException;
                    if (serviceException.ErrorCode == "temporarily_unavailable")
                    {
                        RetryConditionHeaderValue retry = serviceException.Headers.RetryAfter;
                        if (retry.Delta.HasValue)
                        {
                            delay = retry.Delta;
                        }
                        else if (retry.Date.HasValue)
                        {
                            delay = retry.Date.Value.Offset;
                        }
                        if (delay.HasValue)
                        {
                            Thread.Sleep((int)delay.Value.TotalSeconds); // sleep or other
                            keepRetry = true;
                        }
                    }
                    else
                    {
                        throw;
                    }
                }
            } while (keepRetry);

            //token = await authenticationContext.AcquireTokenAsync(resourceUri, cac);
            ResourceTokenLookup[resourceUri] = token;

            log?.Info($"Aquired token which expires {token.ExpiresOn.UtcDateTime}");
            return token.AccessToken;
        }

        public static GraphServiceClient GetGraphClient(GraphEndpoint endPoint = GraphEndpoint.v1)
        {
            var endpointUrl = endPoint == GraphEndpoint.v1 ? "https://graph.microsoft.com/v1.0" : "https://graph.microsoft.com/beta";

            GraphServiceClient client = new GraphServiceClient(endpointUrl, new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    string accessToken = await GetAccessToken(ADALDomain);
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
            return client;
        }

        public static GraphServiceClient GetGraphClientServiceIdentity(TraceWriter log)
        {
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    //var bearerToken = await GetBearerTokenServiceIdentity(log);
                    var bearerToken = await AzureServiceTokenProvider.GetAccessTokenAsync(GraphResourceId);
                    log.Info("Bearer: " + bearerToken);
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                }));
            return client;
        }

        public static async Task<string> GetBearerToken(bool useTenantAdmin = false)
        {
            return await GetAccessToken(ADALDomain, useTenantAdmin);
        }

        public static async Task<string> GetBearerTokenServiceIdentity(TraceWriter log)
        {
            return await AzureServiceTokenProvider.GetAccessTokenAsync(GraphResourceId);
        }

        private static ClientAssertionCertificate GetClientAssertionCertificate()
        {
            var generator = new Certificate.Certificate(_appCert, _appCertKey, "");
            X509Certificate2 cert = generator.GetCertificateFromPEMstring(false);
            ClientAssertionCertificate cac = new ClientAssertionCertificate(_appId, cert);
            return cac;
        }

        public static async Task<ClientContext> GetClientContext(string url, TraceWriter log = null)
        {
            string bearerToken = await GetAccessTokenSharePoint(ADALDomain, url, log);
            var clientContext = new ClientContext(url);
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + bearerToken;
            };
            return clientContext;
        }

    }
}
