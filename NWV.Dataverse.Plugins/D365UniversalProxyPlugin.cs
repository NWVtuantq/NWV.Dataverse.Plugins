using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NWV.Dataverse.Plugins
{
    public class D365UniversalProxyPlugin : IPlugin
    {
        private static readonly HttpClient  _httpClient = new HttpClient();

        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Nhận thông số đầu vào từ PCF / Custom API
                string relativeUrl = context.InputParameters.Contains("RelativeUrl") ? (string)context.InputParameters["RelativeUrl"] : null;
                string httpMethod = context.InputParameters.Contains("HttpMethod") ? (string)context.InputParameters["HttpMethod"] : "GET";
                string requestBody = context.InputParameters.Contains("RequestBody") ? (string)context.InputParameters["RequestBody"] : "";
                string contentType = context.InputParameters.Contains("ContentType") ? (string)context.InputParameters["ContentType"] : "application/json";

                if (string.IsNullOrEmpty(relativeUrl))
                {
                    throw new InvalidPluginExecutionException("Thiếu tham số RelativeUrl.");
                }

                string fnoBaseUrl = GetEnvironmentVariable(service, "nwv_FOBaseUrl");
                string bcTenantName = GetEnvironmentVariable(service, "nwv_BCTenantName");

                if (string.IsNullOrEmpty(fnoBaseUrl))
                {
                    throw new InvalidPluginExecutionException("Chưa cấu hình Environment Variable: nwv_FOBaseUrl.");
                }

                if (string.IsNullOrEmpty(bcTenantName))
                {
                    throw new InvalidPluginExecutionException("Chưa cấu hình Environment Variable: nwv_BCTenantName.");
                }

                // Truyền bcTenantName vào hàm
                string accessToken = GetAccessTokenFromTable(service, tracingService, bcTenantName);

                // Thực thi gọi API bất đồng bộ
                string jsonResponse = Task.Run(() =>
                    CallD365FOAsync(accessToken, fnoBaseUrl, relativeUrl, httpMethod, requestBody, contentType)
                ).GetAwaiter().GetResult();

                // Trả kết quả về cho frontend
                context.OutputParameters["ResponseJson"] = jsonResponse;
            }
            catch (Exception ex)
            {
                tracingService.Trace($"[Lỗi Proxy Chi Tiết] {ex.Message} \n {ex.StackTrace}");
                throw new InvalidPluginExecutionException($"D365 Proxy Error: {ex.Message}");
            }
        }

        private string GetAccessTokenFromTable(IOrganizationService service, ITracingService tracer, string tenantName)
        {
            string tableName = "itl_bctenant";
            string tokenFieldName = "itl_accesstoken";

            // Tên logic của cột Name trong Dataverse
            string nameColumnName = "itl_name";

            string fetchXml = $@"
                <fetch top='1' no-lock='true'>
                    <entity name='{tableName}'>
                    <attribute name='{tokenFieldName}' />
            
                    <filter type='and'>
                        <!-- Lọc theo giá trị lấy từ biến môi trường (vd: 'NWV', 'F&O WMS') -->
                        <condition attribute='{nameColumnName}' operator='eq' value='{tenantName}' />
                
                        <!-- Chỉ lấy Token đang ở trạng thái Active -->
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                </fetch>";

            EntityCollection results = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (results.Entities.Count > 0)
            {
                string token = results.Entities[0].GetAttributeValue<string>(tokenFieldName);
                if (!string.IsNullOrEmpty(token))
                {
                    tracer.Trace($"Đã lấy Access Token thành công từ bảng cho Tenant: {tenantName}.");
                    return token;
                }
            }

            throw new InvalidPluginExecutionException($"Không tìm thấy dữ liệu Access Token đang Active trong bảng {tableName} cho Tenant Name là: '{tenantName}'. Vui lòng kiểm tra lại dữ liệu.");
        }

        private string GetEnvironmentVariable(IOrganizationService service, string variableName)
        {
            string fetchXml = $@"
                <fetch top='1' no-lock='true'>
                  <entity name='environmentvariablevalue'>
                    <attribute name='value' />
                    <link-entity name='environmentvariabledefinition' from='environmentvariabledefinitionid' to='environmentvariabledefinitionid'>
                      <filter>
                        <condition attribute='schemaname' operator='eq' value='{variableName}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

            EntityCollection results = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (results.Entities.Count > 0)
            {
                return results.Entities[0].GetAttributeValue<string>("value");
            }
            return null;
        }

        private async Task<string> CallD365FOAsync(string accessToken, string baseUrl, string relativeUrl, string httpMethod, string body, string contentType)
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            string fullEndpoint = baseUrl + (relativeUrl.StartsWith("/") ? relativeUrl : "/" + relativeUrl);
            HttpResponseMessage fnoRes;

            if (httpMethod.ToUpper() == "POST")
            {
                fnoRes = await _httpClient.PostAsync(fullEndpoint, new StringContent(body, System.Text.Encoding.UTF8, contentType));
            }
            else if (httpMethod.ToUpper() == "PATCH")
            {
                fnoRes = await _httpClient.SendAsync(new HttpRequestMessage(new HttpMethod("PATCH"), fullEndpoint)
                {
                    Content = new StringContent(body, System.Text.Encoding.UTF8, contentType)
                });
            }
            else if (httpMethod.ToUpper() == "PUT")
            {
                fnoRes = await _httpClient.PutAsync(fullEndpoint, new StringContent(body, System.Text.Encoding.UTF8, contentType));
            }
            else if (httpMethod.ToUpper() == "DELETE")
            {
                fnoRes = await _httpClient.DeleteAsync(fullEndpoint);
            }
            else
            {
                fnoRes = await _httpClient.GetAsync(fullEndpoint);
            }

            string responseData = await fnoRes.Content.ReadAsStringAsync();
            if (!fnoRes.IsSuccessStatusCode)
            {
                string errorDetail = string.IsNullOrWhiteSpace(responseData)
                    ? $"HTTP Lỗi {(int)fnoRes.StatusCode} - {fnoRes.ReasonPhrase} (URL: {fullEndpoint})"
                    : responseData;
                throw new Exception(errorDetail);
            }

            return responseData;
        }
    }
}
