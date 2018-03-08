using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json.Linq;

namespace GoogleDriveConverter
{
    public static class Function1
    {
        private static HttpClient client = new HttpClient();

        [FunctionName("Convert_to_Office")]
        public static async Task<HttpResponseMessage> Convert(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post")]HttpRequestMessage req,
            TraceWriter log)
        {
            log.Info("Convert trigger function processed a request.");

            // Get request body
            var data = await req.Content.ReadAsAsync<JObject>();
            string filename = (string)data["name"];

            string mimeType;
            switch ((string)data["type"])
            {
                case "application/vnd.google-apps.spreadsheet":
                    mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    filename += ".xlsx";
                    break;
                case "application/vnd.google-apps.document":
                    mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    filename += ".docx";
                    break;
                case "application/vnd.google-apps.presentation":
                    mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    filename += ".pptx";
                    break;
                default:
                    mimeType = "@DOWNLOAD";
                    break;
            }
            string url;
            if (mimeType.Equals("@DOWNLOAD"))
            {
                url = $"https://content.googleapis.com/drive/v3/files/{Uri.EscapeDataString((string)data["fileId"])}?alt=media";
            }
            else
            {
                url = $"https://content.googleapis.com/drive/v3/files/{Uri.EscapeDataString((string)data["fileId"])}/export?mimeType={Uri.EscapeDataString(mimeType)}"; 
            }
            
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", req.Headers.Authorization.Parameter);
            var response = await client.SendAsync(request);
            response.Headers.Add("filename", filename);
            return response;
        }

        [FunctionName("List_files_in_folder")]
        public static async Task<HttpResponseMessage> List(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post")]HttpRequestMessage req,
            TraceWriter log)
        {
            log.Info("List trigger function processed a request.");
            // Get request body
            var data = await req.Content.ReadAsAsync<JObject>();

            string url = $"https://www.googleapis.com/drive/v3/files?q='{Uri.EscapeDataString((string)data["folderId"])}'+in+parents";
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", req.Headers.Authorization.Parameter);
            var response = await client.SendAsync(request);

            return response;
        }
    }
}
