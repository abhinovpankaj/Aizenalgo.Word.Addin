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
using System.IO;

namespace Aizenalgo.Word.Addin
{
    public partial class ThisAddIn
    {

        public LoginControlHost LoginControl;
        public Microsoft.Office.Tools.CustomTaskPane LoginTaskPane;
        private bool _isUserloggedIn;
        public bool IsUserLoggedIn
        { 
            get { return _isUserloggedIn; }
            set 
            {
                if (value==false)
                {
                    ShowLoginWindow();
                }
                _isUserloggedIn = value;
            }
        }
        private static readonly log4net.ILog log =
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public bool IsActiveDocDocuzen { get; set; }
        public int Mode { get; set; }
        public OfficeRibbon DocuzenRibbon { get; set; }
        public Dictionary<string,DocuzenDocument> DocuzenDocList { get; set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //configurateLogging();
           
            Task.Run(() => {
                
                log.Info("Docuzen Add-in loading.");
                Globals.ThisAddIn.Application.DocumentOpen += Application_DocumentOpen; 
                               
                Globals.ThisAddIn.Application.WindowActivate += Application_WindowActivate;
                DocuzenDocList = new Dictionary<string, DocuzenDocument>();
                log.Info("Docuzen Add-in loaded successfully.");
                UpdateButtonState();
            } );

           
        }

        private void configurateLogging()
        {
            string MyConfigFile = this.GetType().Assembly.ManifestModule.Name + ".config";
            if (File.Exists(MyConfigFile))
            {
                FileInfo fi = new FileInfo(MyConfigFile);
                log4net.Config.XmlConfigurator.Configure(fi);
            }

        }

        public void UpdateButtonState()
        {
            //RibbonButton btn = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnLogout") as RibbonButton;
            //if (btn != null)
            //{
            //    btn.Label = IsUserLoggedIn ? "Logout" : "Login";
            //}
            RibbonButton btnsubmit = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnSubmit") as RibbonButton;
            RibbonButton btnsave = DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Items.FirstOrDefault(s => s.Name == "btnSave") as RibbonButton;

            btnsave.Enabled = IsActiveDocDocuzen;
            btnsubmit.Enabled = IsActiveDocDocuzen;
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
                    UpdateButtonState();
                }
                else
                {
                    IsActiveDocDocuzen = false;
                    DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Visible = false;
                    log.Info($"Docuzen Add-in:{Doc.Name} is a non-docuzen document. Docuzen group will be hidden");
                }
            }
            catch (Exception ex)
            {
                log.Error($"Docuzen Add-in:Failed while activating {Doc.Name}.",ex);
                throw;
            }                           
        }

        public void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
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
                    var logoURL = ReadDocumentProperty(Doc, "logou");
                    docuzendoc.SessionId = sessionId;
                    docuzendoc.UserId = userId;
                    docuzendoc.DocumentId = docId;
                    docuzendoc.LogoURL = logoURL;
                    if (DocuzenDocList.ContainsKey(Doc.Name))
                    {
                        DocuzenDocList.Remove(Doc.Name);
                    }
                    DocuzenDocList.Add(Doc.Name, docuzendoc);
                    IsActiveDocDocuzen = true;
                    DocuzenRibbon.Tabs[0].Groups.FirstOrDefault(x => x.Name == "grpDocuzen").Visible = true;
                }
                else
                    IsActiveDocDocuzen = false;
            }
            catch (Exception ex)
            {
                log.Error($"Docuzen Add-in:Failed while reading docuzen properties in {Doc.Name}.", ex);
                //throw;
            }
            UpdateButtonState();
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

        public void ShowLoginWindow()
        {
            LoginControl control = new LoginControl();
            control.ShowDialog();
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
