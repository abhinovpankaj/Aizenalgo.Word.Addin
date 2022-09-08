using System;
using System.Collections.Generic;
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
        const string VERIFICATIONBASEURL = "http://demo.aizenalgo.com:9016/api/WordProcSessionVerification/WordProcSessionVerification";
        static HttpClient client = new HttpClient();

        public static async Task<ServiceResponse> DocuzenSessionVerification(string sessionId, string docId)
        {
            ServiceResponse verificationResponse=null;
            string endpoint = $"{VERIFICATIONBASEURL}?SessionId={sessionId}&DocID={docId}";
            //endpoint = "http://demo.aizenalgo.com:9016/api/WordProcSessionVerification/WordProcSessionVerification?SessionId=3423sdfdsf45dfvfg332&DocID=3456";
            HttpResponseMessage response = await client.GetAsync(endpoint);
            if (response.IsSuccessStatusCode)
            {
                verificationResponse = await response.Content.ReadAsAsync<ServiceResponse>();
            }
            return verificationResponse;
        }

        public static async Task<ServiceResponse> DocuzenAuthentication(string userName, string password,string sessionId, string docId)
        {
            ServiceResponse verificationResponse = null;
            string endPoint = $"{AUTHENTICATIONBASEURL}?UserName={userName}&" +
                $"Password={password}&SessionId={sessionId}&DocID={docId}";
            //endPoint = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication?UserName=Admin1&Password=Aizant@123&SessionId=123&DocID=1";
            HttpResponseMessage response = await client.GetAsync(endPoint);
            if (response.IsSuccessStatusCode)
            {
                verificationResponse = await response.Content.ReadAsAsync<ServiceResponse>();
            }
            return verificationResponse;
        }
    }
}
