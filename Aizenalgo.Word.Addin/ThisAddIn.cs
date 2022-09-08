using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace Aizenalgo.Word.Addin
{
    public partial class ThisAddIn
    {

        public LoginControlHost LoginControl;
        public Microsoft.Office.Tools.CustomTaskPane LoginTaskPane;
        public bool IsUserLoggedIn { get; set; }
        private static readonly log4net.ILog log =
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public bool IsActiveDocDocuzen { get; set; }

        public OfficeRibbon DocuzenRibbon { get; set; }
        public Dictionary<string,DocuzenDocument> DocuzenDocList { get; set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Task.Run(() => {
                log.Info("Docuzen Add-in loading.");
                Globals.ThisAddIn.Application.DocumentOpen += Application_DocumentOpen;
                //Globals.ThisAddIn.Application. += Application_DocumentOpen;
                //Globals.ThisAddIn.Application.WindowActivate += Application_WindowActivate;
                DocuzenDocList = new Dictionary<string, DocuzenDocument>();
                log.Info("Docuzen Add-in loaded successfully.");
                UpdateButtonLabel();
            } );        
          

        }

        public void UpdateButtonLabel()
        {
            RibbonButton btn = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnLogout") as RibbonButton;
            if (btn != null)
            {
                btn.Label = IsUserLoggedIn ? "Logout" : "Login";
            }
            RibbonButton btnsubmit = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnSubmit") as RibbonButton;
            RibbonButton btnsave = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnSave") as RibbonButton;

            btnsave.Enabled = IsUserLoggedIn;
            btnsubmit.Enabled = IsUserLoggedIn;
        }
        private void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            log.Info($"Docuzen Add-in:{Doc.Name} document activated.");
            try
            {
                if (DocuzenDocList.ContainsKey(Doc.Name))
                {
                    IsActiveDocDocuzen = true;
                    DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Visible = true;
                    log.Info($"Docuzen Add-in:{Doc.Name} is a docuzen document.");

                    //Enable Signin button
                    UpdateButtonLabel();
                }
                else
                {
                    IsActiveDocDocuzen = false;
                    DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Visible = false;
                    log.Info($"Docuzen Add-in:{Doc.Name} is a nn-docuzen document. Docuzen group will be hidden");
                }
            }
            catch (Exception ex)
            {
                log.Error($"Docuzen Add-in:Failed while activating {Doc.Name}.",ex);
                throw;
            }                           
        }

        private void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            log.Info($"Docuzen Add-in:{Doc.Name} document opened. Reading for dozuzen doc properties");
            try
            {
                DocuzenDocument docuzendoc = new DocuzenDocument();
                Microsoft.Office.Core.DocumentProperties properties;
                properties = (Office.DocumentProperties)Doc.CustomDocumentProperties;

                var docId = ReadDocumentProperty(Doc, "DVId");
                if (docId != null)
                {
                    var sessionId = ReadDocumentProperty(Doc, "SToken");
                    var userId = ReadDocumentProperty(Doc, "Uid");
                    docuzendoc.SessionId = sessionId;
                    docuzendoc.UserId = userId;
                    docuzendoc.DocumentId = docId;
                    DocuzenDocList.Add(Doc.Name, docuzendoc);
                }
            }
            catch (Exception ex)
            {
                log.Error($"Docuzen Add-in:Failed while reading docuzen properties in {Doc.Name}.", ex);
                //throw;
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
