namespace HurTest
{
    partial class HurTestRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HurTestRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HurTestRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.HurTestDesignGroup = this.Factory.CreateRibbonGroup();
            this.btnInsertGroup = this.Factory.CreateRibbonButton();
            this.btnInsertQuestion = this.Factory.CreateRibbonButton();
            this.btnInsertChoice = this.Factory.CreateRibbonButton();
            this.btnInsertNumeric = this.Factory.CreateRibbonButton();
            this.btnInsertList = this.Factory.CreateRibbonButton();
            this.HurTestOutputGroup = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.HurTestDesignGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.HurTestDesignGroup);
            this.tab1.Groups.Add(this.HurTestOutputGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // HurTestDesignGroup
            // 
            this.HurTestDesignGroup.Items.Add(this.btnInsertGroup);
            this.HurTestDesignGroup.Items.Add(this.btnInsertQuestion);
            this.HurTestDesignGroup.Items.Add(this.btnInsertChoice);
            this.HurTestDesignGroup.Items.Add(this.btnInsertNumeric);
            this.HurTestDesignGroup.Items.Add(this.btnInsertList);
            this.HurTestDesignGroup.Label = "HurTest Tasarım";
            this.HurTestDesignGroup.Name = "HurTestDesignGroup";
            // 
            // btnInsertGroup
            // 
            this.btnInsertGroup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertGroup.Image = global::HurTest.Properties.Resources.btnImage_InsertGroup;
            this.btnInsertGroup.Label = "Soru Grubu Ekle";
            this.btnInsertGroup.Name = "btnInsertGroup";
            this.btnInsertGroup.ShowImage = true;
            this.btnInsertGroup.SuperTip = "Soru Grubu Ekle";
            this.btnInsertGroup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertGroup_Click);
            // 
            // btnInsertQuestion
            // 
            this.btnInsertQuestion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertQuestion.Image = global::HurTest.Properties.Resources.btnImage_InsertQuestion;
            this.btnInsertQuestion.Label = "Soru Ekle";
            this.btnInsertQuestion.Name = "btnInsertQuestion";
            this.btnInsertQuestion.ShowImage = true;
            this.btnInsertQuestion.SuperTip = "Soru Ekle";
            this.btnInsertQuestion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertQuestion_Click);
            // 
            // btnInsertChoice
            // 
            this.btnInsertChoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertChoice.Image = global::HurTest.Properties.Resources.btnImage_InsertChoice;
            this.btnInsertChoice.Label = "Seçenek Ekle";
            this.btnInsertChoice.Name = "btnInsertChoice";
            this.btnInsertChoice.ShowImage = true;
            this.btnInsertChoice.SuperTip = "Seçenek Ekle";
            this.btnInsertChoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertChoice_Click);
            // 
            // btnInsertNumeric
            // 
            this.btnInsertNumeric.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertNumeric.Image = ((System.Drawing.Image)(resources.GetObject("btnInsertNumeric.Image")));
            this.btnInsertNumeric.Label = "Sayısal Alan Ekle";
            this.btnInsertNumeric.Name = "btnInsertNumeric";
            this.btnInsertNumeric.ShowImage = true;
            this.btnInsertNumeric.SuperTip = "Sayısal Alan Ekle";
            this.btnInsertNumeric.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertNumeric_Click);
            // 
            // btnInsertList
            // 
            this.btnInsertList.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertList.Image = global::HurTest.Properties.Resources.btnImage_InsertOptions;
            this.btnInsertList.Label = "Liste Alanı Ekle";
            this.btnInsertList.Name = "btnInsertList";
            this.btnInsertList.ShowImage = true;
            this.btnInsertList.SuperTip = "Liste Alanı Ekle";
            // 
            // HurTestOutputGroup
            // 
            this.HurTestOutputGroup.Label = "HurTest Çıktı";
            this.HurTestOutputGroup.Name = "HurTestOutputGroup";
            // 
            // HurTestRibbon
            // 
            this.Name = "HurTestRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HurTestRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.HurTestDesignGroup.ResumeLayout(false);
            this.HurTestDesignGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup HurTestDesignGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertChoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertNumeric;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertList;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup HurTestOutputGroup;
    }

    partial class ThisRibbonCollection
    {
        internal HurTestRibbon HurTestRibbon
        {
            get { return this.GetRibbon<HurTestRibbon>(); }
        }
    }
}
