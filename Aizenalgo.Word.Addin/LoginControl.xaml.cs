using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Aizenalgo.Word.Addin
{
    /// <summary>
    /// Interaction logic for LoginControl.xaml
    /// </summary>
    public partial class LoginControl : UserControl
    {
        private static readonly log4net.ILog log =
                        log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public LoginControl()
        {
            InitializeComponent();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            //call service and login
            string activeDocName = Globals.ThisAddIn.Application.ActiveDocument.Name;
            var activeDocuzen = Globals.ThisAddIn.DocuzenDocList[activeDocName];
            string userName = this.username.Text;
            string password = this.password.Password;
            log.Info("SignIn button Clicked");
            if (activeDocuzen!=null)
            {
                log.Info("Docuzen doc found.");
                ServiceResponse response = await DocuzenService.DocuzenAuthentication(userName, password, activeDocuzen.SessionId, activeDocuzen.DocumentId);
                if (response.MsgType == "Success")
                {
                    //close the pane.
                    Globals.ThisAddIn.LoginTaskPane.Visible = false;
                    Globals.ThisAddIn.IsUserLoggedIn = true;
                    log.Info("Loggedin successfully.");
                }
                else
                {
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
    }
}
