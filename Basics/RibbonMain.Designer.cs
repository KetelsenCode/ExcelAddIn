namespace Basics
{
    partial class RibbonMain : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMain()
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
            this.btn1 = this.Factory.CreateRibbonButton();
            this.btn2 = this.Factory.CreateRibbonButton();
            this.lb1 = this.Factory.CreateRibbonLabel();
            this.eb1 = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn1);
            this.group1.Items.Add(this.btn2);
            this.group1.Items.Add(this.lb1);
            this.group1.Items.Add(this.eb1);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btn1
            // 
            this.btn1.Label = "insert";
            this.btn1.Name = "btn1";
            this.btn1.Click += Btn1_Click;
            // 
            // btn2
            // 
            this.btn2.Label = "read";
            this.btn2.Name = "btn2";
            this.btn2.Click += Btn2_Click;
            // 
            // lb1
            // 
            this.lb1.Label = "";
            this.lb1.Name = "lb1";
            // 
            // eb1
            // 
            this.eb1.Label = "";
            this.eb1.Name = "eb1";
            this.eb1.Text = null;
            // 
            // RibbonMain
            // 
            this.Name = "RibbonMain";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMain_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }





        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn2;

        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lb1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox eb1;

    }

    partial class ThisRibbonCollection
    {
        internal RibbonMain RibbonMain
        {
            get { return this.GetRibbon<RibbonMain>(); }
        }
    }
}
