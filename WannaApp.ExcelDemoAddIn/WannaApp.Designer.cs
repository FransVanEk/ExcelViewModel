namespace WannaApp.ExcelDemoAddIn
{
    partial class WannaApp : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WannaApp()
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
            this.WannaAppTabFile = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_LoadListObject = this.Factory.CreateRibbonButton();
            this.Objects = this.Factory.CreateRibbonGroup();
            this.LoadObjects = this.Factory.CreateRibbonButton();
            this.GetObjects = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_loadDemo = this.Factory.CreateRibbonButton();
            this.WannaAppTabFile.SuspendLayout();
            this.group1.SuspendLayout();
            this.Objects.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // WannaAppTabFile
            // 
            this.WannaAppTabFile.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.WannaAppTabFile.Groups.Add(this.group1);
            this.WannaAppTabFile.Groups.Add(this.Objects);
            this.WannaAppTabFile.Groups.Add(this.group2);
            this.WannaAppTabFile.Label = "Wanna App Tab File";
            this.WannaAppTabFile.Name = "WannaAppTabFile";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_LoadListObject);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btn_LoadListObject
            // 
            this.btn_LoadListObject.Label = "Load";
            this.btn_LoadListObject.Name = "btn_LoadListObject";
            this.btn_LoadListObject.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LoadListObject_Click);
            // 
            // Objects
            // 
            this.Objects.Items.Add(this.LoadObjects);
            this.Objects.Items.Add(this.GetObjects);
            this.Objects.Label = "group2";
            this.Objects.Name = "Objects";
            // 
            // LoadObjects
            // 
            this.LoadObjects.Label = "Load objects";
            this.LoadObjects.Name = "LoadObjects";
            this.LoadObjects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoadObjects_Click);
            // 
            // GetObjects
            // 
            this.GetObjects.Label = "Get Objects";
            this.GetObjects.Name = "GetObjects";
            this.GetObjects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetObjects_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_loadDemo);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btn_loadDemo
            // 
            this.btn_loadDemo.Label = "Load Demo objecten";
            this.btn_loadDemo.Name = "btn_loadDemo";
            this.btn_loadDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_loadDemo_Click);
            // 
            // WannaApp
            // 
            this.Name = "WannaApp";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.WannaAppTabFile);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WannaApp_Load);
            this.WannaAppTabFile.ResumeLayout(false);
            this.WannaAppTabFile.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Objects.ResumeLayout(false);
            this.Objects.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab WannaAppTabFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LoadListObject;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Objects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LoadObjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetObjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_loadDemo;
    }

    partial class ThisRibbonCollection
    {
        internal WannaApp WannaApp
        {
            get { return this.GetRibbon<WannaApp>(); }
        }
    }
}
