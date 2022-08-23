namespace Aizenalgo.Word.Addin
{
    partial class Aizenalgo : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Aizenalgo()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpDocuzen = this.Factory.CreateRibbonGroup();
            this.btnSubmit = this.Factory.CreateRibbonButton();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpDocuzen.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpDocuzen);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
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
            this.btnSubmit.ShowImage = true;
            // 
            // btnSave
            // 
            this.btnSave.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSave.Label = "Save";
            this.btnSave.Name = "btnSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // Aizenalgo
            // 
            this.Name = "Aizenalgo";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Aizenalgo_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpDocuzen.ResumeLayout(false);
            this.grpDocuzen.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDocuzen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubmit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
    }

    partial class ThisRibbonCollection
    {
        internal Aizenalgo Aizenalgo
        {
            get { return this.GetRibbon<Aizenalgo>(); }
        }
    }
}
