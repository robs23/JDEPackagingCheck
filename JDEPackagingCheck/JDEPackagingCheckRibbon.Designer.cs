
namespace JDEPackagingCheck
{
    partial class JDEPackagingCheckRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public JDEPackagingCheckRibbon()
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
            this.tabJDEPackaging = this.Factory.CreateRibbonTab();
            this.grpScheduleKontrol = this.Factory.CreateRibbonGroup();
            this.btnShowCoverage = this.Factory.CreateRibbonButton();
            this.btnHideCoverage = this.Factory.CreateRibbonButton();
            this.grpImport = this.Factory.CreateRibbonGroup();
            this.btnImportInventories = this.Factory.CreateRibbonButton();
            this.btnImportDeliveries = this.Factory.CreateRibbonButton();
            this.tabJDEPackaging.SuspendLayout();
            this.grpScheduleKontrol.SuspendLayout();
            this.grpImport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabJDEPackaging
            // 
            this.tabJDEPackaging.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabJDEPackaging.Groups.Add(this.grpScheduleKontrol);
            this.tabJDEPackaging.Groups.Add(this.grpImport);
            this.tabJDEPackaging.Label = "JDE Opakowania";
            this.tabJDEPackaging.Name = "tabJDEPackaging";
            // 
            // grpScheduleKontrol
            // 
            this.grpScheduleKontrol.Items.Add(this.btnShowCoverage);
            this.grpScheduleKontrol.Items.Add(this.btnHideCoverage);
            this.grpScheduleKontrol.Label = "Kontrola harmonogramu";
            this.grpScheduleKontrol.Name = "grpScheduleKontrol";
            // 
            // btnShowCoverage
            // 
            this.btnShowCoverage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnShowCoverage.Image = global::JDEPackagingCheck.Properties.Resources.colorizeRows;
            this.btnShowCoverage.Label = "Pokaż pokrycie";
            this.btnShowCoverage.Name = "btnShowCoverage";
            this.btnShowCoverage.ShowImage = true;
            this.btnShowCoverage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowCoverage_Click);
            // 
            // btnHideCoverage
            // 
            this.btnHideCoverage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHideCoverage.Image = global::JDEPackagingCheck.Properties.Resources.hideCoverage;
            this.btnHideCoverage.Label = "Ukryj pokrycie";
            this.btnHideCoverage.Name = "btnHideCoverage";
            this.btnHideCoverage.ShowImage = true;
            this.btnHideCoverage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHideCoverage_Click);
            // 
            // grpImport
            // 
            this.grpImport.Items.Add(this.btnImportInventories);
            this.grpImport.Items.Add(this.btnImportDeliveries);
            this.grpImport.Label = "Import";
            this.grpImport.Name = "grpImport";
            // 
            // btnImportInventories
            // 
            this.btnImportInventories.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportInventories.Image = global::JDEPackagingCheck.Properties.Resources.inventory;
            this.btnImportInventories.Label = "Importuj zapasy";
            this.btnImportInventories.Name = "btnImportInventories";
            this.btnImportInventories.ShowImage = true;
            this.btnImportInventories.SuperTip = "Importuj stany magazynowe pobrane z SAPa";
            this.btnImportInventories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportInventories_Click);
            // 
            // btnImportDeliveries
            // 
            this.btnImportDeliveries.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportDeliveries.Image = global::JDEPackagingCheck.Properties.Resources.delivery;
            this.btnImportDeliveries.Label = "Importuj dostawy";
            this.btnImportDeliveries.Name = "btnImportDeliveries";
            this.btnImportDeliveries.ShowImage = true;
            this.btnImportDeliveries.SuperTip = "Importuj dostawy komponentów pobrane z SAPa";
            // 
            // JDEPackagingCheckRibbon
            // 
            this.Name = "JDEPackagingCheckRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabJDEPackaging);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.JDEPackagingCheckRibbon_Load);
            this.tabJDEPackaging.ResumeLayout(false);
            this.tabJDEPackaging.PerformLayout();
            this.grpScheduleKontrol.ResumeLayout(false);
            this.grpScheduleKontrol.PerformLayout();
            this.grpImport.ResumeLayout(false);
            this.grpImport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabJDEPackaging;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpScheduleKontrol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowCoverage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHideCoverage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportInventories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportDeliveries;
    }

    partial class ThisRibbonCollection
    {
        internal JDEPackagingCheckRibbon JDEPackagingCheckRibbon
        {
            get { return this.GetRibbon<JDEPackagingCheckRibbon>(); }
        }
    }
}
