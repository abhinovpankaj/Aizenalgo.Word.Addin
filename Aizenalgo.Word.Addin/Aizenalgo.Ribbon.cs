using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aizenalgo.Word.Addin
{
    public partial class Aizenalgo
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

        }
    }
}
