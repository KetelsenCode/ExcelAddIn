namespace Advanced
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
            this.tab1 = Factory.CreateRibbonTab();
            this.group1 = Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.btn1 = Factory.CreateRibbonButton();
            this.btn2 = Factory.CreateRibbonButton();
            this.btn3 = Factory.CreateRibbonButton();
            this.lb1 = Factory.CreateRibbonLabel();
            this.eb1 = Factory.CreateRibbonEditBox();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Name = "tab1";
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabAddIns";
            this.tab1.Groups.Add(this.group1);
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            this.group1.Items.Add(btn1);
            this.group1.Items.Add(btn2);
            this.group1.Items.Add(btn3);
            this.group1.Items.Add(lb1);
            this.group1.Items.Add(eb1);
            // 
            // btn1
            // 
            this.btn1.Label = "insert";
            this.btn1.Name = "btn1";
            this.btn1.Click += Btn1_Click;
            // 
            // btn2
            // 
            this.btn2.Label = "Calculate rows";
            this.btn2.Name = "btn2";
            this.btn2.Click += Btn2_Click;
            // 
            // btn3
            // 
            this.btn3.Label = "insert";
            this.btn3.Name = "btn3";
            this.btn3.Click += Btn3_Click;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn3;

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
