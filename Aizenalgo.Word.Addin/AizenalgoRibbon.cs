using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
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
            //Globals.ThisAddIn.Application.DocumentOpen += Globals.ThisAddIn.Application_DocumentOpen;
        }

        private async void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.StatusBar = "Saving to Aizenalgo server...";
            Globals.ThisAddIn.Mode = 1;
            string activeDocName = Globals.ThisAddIn.Application.ActiveDocument.Name;
            var activeDocuzen = Globals.ThisAddIn.DocuzenDocList[activeDocName];
            string tempPath = System.IO.Path.GetTempPath();
            string fPath = Path.Combine(tempPath, activeDocName);
            File.Copy(Path.Combine(Globals.ThisAddIn.Application.ActiveDocument.Path, activeDocName), fPath, true);
            try
            {
                log.Info("Submit button Clicked");
                if (activeDocuzen != null)
                {
                    log.Info("Docuzen doc found.");
                    await Dispatcher.CurrentDispatcher.Invoke(async () =>
                    {
                        ServiceResponse response = await DocuzenService.DocuzenSessionVerification(activeDocuzen.SessionId, activeDocuzen.DocumentId, fPath, activeDocName, 1);
                        if (response != null)
                        {
                            if (response.MsgType == "Success")
                            {
                                //close the pane.
                                MessageBox.Show("Document Saved successfully");
                                log.Info("Session verified successfully.Submssion will start.");
                            }
                            else
                            {
                                Globals.ThisAddIn.Application.StatusBar = "Error from Server " + response.MsgError;
                                // Dispatcher.CurrentDispatcher.Invoke(() => ShowLoginWindow());
                                Globals.ThisAddIn.IsUserLoggedIn = false;
                                //Globals.ThisAddIn.ShowLoginWindow();

                                log.Info("Log-in failed.");
                            }
                        }
                        else
                        {
                            //Dispatcher.CurrentDispatcher.Invoke(() => ShowLoginWindow());
                            //Globals.ThisAddIn.ShowLoginWindow();
                            Globals.ThisAddIn.Application.StatusBar = "Error from Server: No response recieved. ";
                            Globals.ThisAddIn.IsUserLoggedIn = false;
                            log.Info("Log-in failed.");
                        }
                    });


                }
                else
                {
                    //log
                    log.Info("No Docuzen doc found in dictionary.");
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception occured while saving document"+ ex.Message);
                MessageBox.Show("Exception while saving the document.");
            }
            
            Globals.ThisAddIn.UpdateButtonState();
            Globals.ThisAddIn.Application.StatusBar = "";
        }

        [STAThread]
        private async void btnSubmit_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.StatusBar = "Submitting to Aizenalgo server...";
            Globals.ThisAddIn.Mode = 2;
            //call service
            string activeDocName = Globals.ThisAddIn.Application.ActiveDocument.Name;
            var activeDocuzen = Globals.ThisAddIn.DocuzenDocList[activeDocName];
            string tempPath = System.IO.Path.GetTempPath();
            string fPath = Path.Combine(tempPath,activeDocName);
            File.Copy(Path.Combine(Globals.ThisAddIn.Application.ActiveDocument.Path, activeDocName), fPath,true);
            try
            {
                if (activeDocuzen != null)
                {
                    log.Info("Docuzen doc found.");
                    await Dispatcher.CurrentDispatcher.Invoke(async () =>
                    {
                        ServiceResponse response = await DocuzenService.DocuzenSessionVerification(activeDocuzen.SessionId, activeDocuzen.DocumentId, fPath, activeDocName, 2);
                        if (response != null)
                        {
                            if (response.MsgType == "Success")
                            {
                                //close the pane.
                                MessageBox.Show("Document Submitted successfully");
                                log.Info("Session verified successfully.Submssion will start.");
                            }
                            else
                            {
                                Globals.ThisAddIn.Application.StatusBar = "Error from Server " + response.MsgError;
                                // Dispatcher.CurrentDispatcher.Invoke(() => ShowLoginWindow());
                                Globals.ThisAddIn.IsUserLoggedIn = false;
                                //Globals.ThisAddIn.ShowLoginWindow();

                                log.Info("Log-in failed.");
                            }
                        }
                        else
                        {
                            Globals.ThisAddIn.Application.StatusBar = "Error from Server: No response recieved. ";
                            //Dispatcher.CurrentDispatcher.Invoke(() => ShowLoginWindow());
                            //Globals.ThisAddIn.ShowLoginWindow();
                            Globals.ThisAddIn.IsUserLoggedIn = false;
                            log.Info("Log-in failed.");
                        }
                    });


                }
                else
                {
                    //log
                    log.Info("No Docuzen doc found in dictionary.");
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception occured while submitting document" + ex.Message);
                MessageBox.Show("Exception while submitting the document.");
                throw;
            }
            log.Info("Submit button Clicked");
            
            Globals.ThisAddIn.UpdateButtonState();
            Globals.ThisAddIn.Application.StatusBar = "";
        }
        
        void ShowLoginWindow()
        {
            LoginControl control = new LoginControl();
            control.ShowDialog();
            //if (Globals.ThisAddIn.LoginTaskPane == null)
            //{

            //    log.Info("Initializing Docuzen login panel.");
            //    Globals.ThisAddIn.LoginTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new LoginControlHost(), "Docuzen Login");
            //    Globals.ThisAddIn.LoginTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            //    Globals.ThisAddIn.LoginTaskPane.Width = 600;
            //    Globals.ThisAddIn.LoginTaskPane.Visible = true;
            //}
            //else
            //    Globals.ThisAddIn.LoginTaskPane.Visible = true;
        }
    }

    
}
