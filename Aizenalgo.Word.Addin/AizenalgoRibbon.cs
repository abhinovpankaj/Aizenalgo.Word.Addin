using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;

namespace Aizenalgo.Word.Addin
{
    public partial class AizenalgoRibbon
    {
        private static readonly log4net.ILog log =
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private void Aizenalgo_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.DocuzenRibbon = sender as OfficeRibbon;
        }

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLogout_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton btn = sender as RibbonButton;
            if (btn!=null)
            {
                if (btn.Label=="Login")
                {
                    //Show Login Propmt.
                    Dispatcher.CurrentDispatcher.Invoke(()=>showhidePanel());
                    
                }
                else
                {
                    Globals.ThisAddIn.IsUserLoggedIn = false;
                    
                }
            }
            Globals.ThisAddIn.UpdateButtonLabel();
        }

        private async void btnSubmit_Click(object sender, RibbonControlEventArgs e)
        {
            //call service
            string activeDocName = Globals.ThisAddIn.Application.ActiveDocument.Name;
            var activeDocuzen = Globals.ThisAddIn.DocuzenDocList[activeDocName];
            
            log.Info("Submit button Clicked");
            if (activeDocuzen != null)
            {
                log.Info("Docuzen doc found.");
                ServiceResponse response = await DocuzenService.DocuzenSessionVerification(activeDocuzen.SessionId, activeDocuzen.DocumentId);
                if (response.MsgType == "Success")
                {
                    //close the pane.
                    
                    log.Info("Session verified successfully.Submssion will start.");
                }
                else
                {
                    Dispatcher.CurrentDispatcher.Invoke(() => showhidePanel());
                    Globals.ThisAddIn.LoginTaskPane.Visible = true;
                    Globals.ThisAddIn.IsUserLoggedIn = false;
                    log.Info("Log-in failed.");
                }
            }
            else
            {
                //log
                log.Info("No Docuzen doc found in dictionary.");
            }
            Globals.ThisAddIn.UpdateButtonLabel();
        }
        void showhidePanel()
        {
            if (Globals.ThisAddIn.LoginTaskPane == null)
            {

                log.Info("Initializing Docuzen login panel.");
                Globals.ThisAddIn.LoginTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new LoginControlHost(), "Docuzen Login");
                Globals.ThisAddIn.LoginTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                Globals.ThisAddIn.LoginTaskPane.Width = 600;
                Globals.ThisAddIn.LoginTaskPane.Visible = true;
            }
            else
                Globals.ThisAddIn.LoginTaskPane.Visible = true;
        }
    }

    
}
