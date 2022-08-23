using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;

namespace Aizenalgo.Word.Addin
{
    public partial class ThisAddIn
    {
        public bool IsActiveDocDocuzen { get; set; }        

        public Dictionary<string,DocuzenDocument> DocuzenDocList { get; set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.DocumentOpen += Application_DocumentOpen;
            Globals.ThisAddIn.Application.WindowActivate += Application_WindowActivate; ;
        }

        private void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            if (DocuzenDocList.ContainsKey(Doc.Name))
            {
                IsActiveDocDocuzen = true;
            }
            else
                IsActiveDocDocuzen = false;
        }

        private void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            DocuzenDocument docuzendoc = new DocuzenDocument();
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)Doc.CustomDocumentProperties;

            var docId = ReadDocumentProperty(Doc, "DocumentId");
            if (docId!=null)
            {
                var sessionId = ReadDocumentProperty(Doc, "DocumentId"); 
                var userId = ReadDocumentProperty(Doc, "DocumentId"); 
                docuzendoc.SessionId = sessionId;
                docuzendoc.UserId = userId;
                docuzendoc.DocumentId = docId;
                DocuzenDocList.Add(Doc.Name,docuzendoc);
            }           
        }
        private string ReadDocumentProperty(Microsoft.Office.Interop.Word.Document Doc,string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)Doc.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
