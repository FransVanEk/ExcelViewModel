namespace WannaApp.Excel.DemoAdd_in
{
    partial class DemoAdd_inRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DemoAdd_inRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_BasicUsage_LoadIntoExcel = this.Factory.CreateRibbonButton();
            this.Objects = this.Factory.CreateRibbonGroup();
            this.btn_LoadObjectDataIntoExcel = this.Factory.CreateRibbonButton();
            this.Validations = this.Factory.CreateRibbonGroup();
            this.btn_LoadValidations = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.Objects.SuspendLayout();
            this.Validations.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.Objects);
            this.tab1.Groups.Add(this.Validations);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_BasicUsage_LoadIntoExcel);
            this.group1.Label = "Basic usage";
            this.group1.Name = "group1";
            // 
            // btn_BasicUsage_LoadIntoExcel
            // 
            this.btn_BasicUsage_LoadIntoExcel.Label = "Load data into excel";
            this.btn_BasicUsage_LoadIntoExcel.Name = "btn_BasicUsage_LoadIntoExcel";
            this.btn_BasicUsage_LoadIntoExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_BasicUsage_LoadIntoExcel_Click);
            // 
            // Objects
            // 
            this.Objects.Items.Add(this.btn_LoadObjectDataIntoExcel);
            this.Objects.Label = "Objects";
            this.Objects.Name = "Objects";
            // 
            // btn_LoadObjectDataIntoExcel
            // 
            this.btn_LoadObjectDataIntoExcel.Label = "Load Object data into excel";
            this.btn_LoadObjectDataIntoExcel.Name = "btn_LoadObjectDataIntoExcel";
            this.btn_LoadObjectDataIntoExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadObjectDataIntoExcel_Click);
            // 
            // Validations
            // 
            this.Validations.Items.Add(this.btn_LoadValidations);
            this.Validations.Label = "AddValidations";
            this.Validations.Name = "Validations";
            // 
            // btn_LoadValidations
            // 
            this.btn_LoadValidations.Label = "Load Validations";
            this.btn_LoadValidations.Name = "btn_LoadValidations";
            this.btn_LoadValidations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadValidations_Click);
            // 
            // DemoAdd_inRibbon
            // 
            this.Name = "DemoAdd_inRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DemoAdd_inRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Objects.ResumeLayout(false);
            this.Objects.PerformLayout();
            this.Validations.ResumeLayout(false);
            this.Validations.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_BasicUsage_LoadIntoExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Objects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LoadObjectDataIntoExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Validations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LoadValidations;
    }

    partial class ThisRibbonCollection
    {
        internal DemoAdd_inRibbon DemoAdd_inRibbon
        {
            get { return this.GetRibbon<DemoAdd_inRibbon>(); }
        }
    }
}
