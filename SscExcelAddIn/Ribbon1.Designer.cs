#pragma warning disable 1591
using System.Drawing;

namespace SscExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.editSheetGroup = this.Factory.CreateRibbonGroup();
            this.ReplaceButton = this.Factory.CreateRibbonButton();
            this.ZebraButton = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.etcGroup = this.Factory.CreateRibbonGroup();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.TestControlButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.editSheetGroup.SuspendLayout();
            this.etcGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.editSheetGroup);
            this.tab1.Groups.Add(this.etcGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // editSheetGroup
            // 
            this.editSheetGroup.Items.Add(this.ReplaceButton);
            this.editSheetGroup.Items.Add(this.ZebraButton);
            this.editSheetGroup.Items.Add(this.button1);
            this.editSheetGroup.Label = "シート編集";
            this.editSheetGroup.Name = "editSheetGroup";
            // 
            // ReplaceButton
            // 
            this.ReplaceButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReplaceButton.Image = global::SscExcelAddIn.Properties.Resources.regex_icon_132036;
            this.ReplaceButton.Label = "高度な\n置換";
            this.ReplaceButton.Name = "ReplaceButton";
            this.ReplaceButton.ShowImage = true;
            this.ReplaceButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplaceButton_Click);
            // 
            // ZebraButton
            // 
            this.ZebraButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ZebraButton.Image = global::SscExcelAddIn.Properties.Resources.zebra;
            this.ZebraButton.Label = "行スキップ\n選択";
            this.ZebraButton.Name = "ZebraButton";
            this.ZebraButton.ShowImage = true;
            this.ZebraButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SkipSelectButton_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // etcGroup
            // 
            this.etcGroup.Items.Add(this.AboutButton);
            this.etcGroup.Items.Add(this.TestControlButton);
            this.etcGroup.Label = "etc";
            this.etcGroup.Name = "etcGroup";
            // 
            // AboutButton
            // 
            this.AboutButton.Label = "クレジット";
            this.AboutButton.Name = "AboutButton";
            this.AboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);
            // 
            // TestControlButton
            // 
            this.TestControlButton.Label = "TestControl";
            this.TestControlButton.Name = "TestControlButton";
            this.TestControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestControlButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.editSheetGroup.ResumeLayout(false);
            this.editSheetGroup.PerformLayout();
            this.etcGroup.ResumeLayout(false);
            this.etcGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup editSheetGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReplaceButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ZebraButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup etcGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestControlButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
