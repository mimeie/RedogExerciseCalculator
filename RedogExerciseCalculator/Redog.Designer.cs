namespace RedogExerciseCalculator
{
    partial class Redog : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Redog()
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
            this.Initialize = this.Factory.CreateRibbonButton();
            this.Calculate = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Initialize);
            this.group1.Items.Add(this.Calculate);
            this.group1.Label = "Redog Übungen";
            this.group1.Name = "group1";
            // 
            // Initialize
            // 
            this.Initialize.Image = global::RedogExerciseCalculator.Properties.Resources.refresh;
            this.Initialize.Label = "Initialisieren";
            this.Initialize.Name = "Initialize";
            this.Initialize.ShowImage = true;
            this.Initialize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Initialize_Click);
            // 
            // Calculate
            // 
            this.Calculate.Image = global::RedogExerciseCalculator.Properties.Resources.RedogHund;
            this.Calculate.Label = "Berechnen";
            this.Calculate.Name = "Calculate";
            this.Calculate.ShowImage = true;
            this.Calculate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Calculate_Click);
            // 
            // Redog
            // 
            this.Name = "Redog";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Redog_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Initialize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Calculate;
    }

    partial class ThisRibbonCollection
    {
        internal Redog Redog
        {
            get { return this.GetRibbon<Redog>(); }
        }
    }
}
