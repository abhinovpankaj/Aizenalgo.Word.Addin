namespace Aizenalgo.Word.Addin
{
    partial class AizenalgoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AizenalgoRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.docuzenTab = this.Factory.CreateRibbonTab();
            this.grpDocuzen = this.Factory.CreateRibbonGroup();
            this.btnSubmit = this.Factory.CreateRibbonButton();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.docuzenTab.SuspendLayout();
            this.grpDocuzen.SuspendLayout();
            this.SuspendLayout();
            // 
            // docuzenTab
            // 
            this.docuzenTab.Groups.Add(this.grpDocuzen);
            this.docuzenTab.Label = "Docuzen";
            this.docuzenTab.Name = "docuzenTab";
            // 
            // grpDocuzen
            // 
            this.grpDocuzen.Items.Add(this.btnSubmit);
            this.grpDocuzen.Items.Add(this.btnSave);
            this.grpDocuzen.Label = "Docuzen";
            this.grpDocuzen.Name = "grpDocuzen";
            // 
            // btnSubmit
            // 
            this.btnSubmit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSubmit.Label = "Submit";
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.OfficeImageId = "SendUpdate";
            this.btnSubmit.ShowImage = true;
            this.btnSubmit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSubmit_Click);
            // 
            // btnSave
            // 
            this.btnSave.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSave.Label = "Save";
            this.btnSave.Name = "btnSave";
            this.btnSave.OfficeImageId = "FileSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // AizenalgoRibbon
            // 
            this.Name = "AizenalgoRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.docuzenTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Aizenalgo_Load);
            this.docuzenTab.ResumeLayout(false);
            this.docuzenTab.PerformLayout();
            this.grpDocuzen.ResumeLayout(false);
            this.grpDocuzen.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab docuzenTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDocuzen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubmit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
    }

    partial class ThisRibbonCollection
    {
        internal AizenalgoRibbon Aizenalgo
        {
            get { return this.GetRibbon<AizenalgoRibbon>(); }
        }
    }
}
