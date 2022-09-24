using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aizenalgo.Word.Addin
{
    public class DocuzenService
    {
        const string AUTHENTICATIONBASEURL = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
        const string VERIFICATIONBASEURL = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";
        static HttpClient client = new HttpClient();

        public static async Task<ServiceResponse> DocuzenSessionVerification(string sessionId, string docId,string filePath, string fileName,int type)
        {
            ServiceResponse verificationResponse = null;
            string endpoint = $"{VERIFICATIONBASEURL}?SessionId={sessionId}&DocID={docId}&Mode={type}";
            //endpoint = "http://demo.aizenalgo.com:9016/api/WordProcSessionVerification/WordProcSessionVerification?SessionId=3423sdfdsf45dfvfg332&DocID=3456";

            using (var formData = new MultipartFormDataContent())
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[fs.Length];
                    int bytesRead = fs.Read(buffer, 0, buffer.Length);
                    formData.Add(new ByteArrayContent(buffer), "file",fileName);
                    HttpResponseMessage response = await client.PostAsync(endpoint,formData);
                    if (response.IsSuccessStatusCode)
                    {
                        verificationResponse = await response.Content.ReadAsAsync<ServiceResponse>();
                    }
                }
            }
            return verificationResponse;
        }

        public static async Task<ServiceResponse> DocuzenAuthentication(string userName, string password,string sessionId, string docId, string filePath, string fileName,int type)
        {
            ServiceResponse verificationResponse = null;
            string endPoint = $"{AUTHENTICATIONBASEURL}?UserName={userName}&" +
                $"Password={password}&SessionId={sessionId}&DocID={docId}&Mode={type}";
            //endPoint = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication?UserName=Admin1&Password=Aizant@123&SessionId=123&DocID=1";
            
            using (var formData = new MultipartFormDataContent())
            {
                
                formData.Add(new ByteArrayContent(File.ReadAllBytes(filePath)), "file", fileName);
                HttpResponseMessage response = await client.PostAsync(endPoint, formData);
                if (response.IsSuccessStatusCode)
                {
                    verificationResponse = await response.Content.ReadAsAsync<ServiceResponse>();                                      
                }
            }
            
            return verificationResponse;
        }
    }
}
