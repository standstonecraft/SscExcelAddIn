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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.updateGroup = this.Factory.CreateRibbonGroup();
            this.editSheetGroup = this.Factory.CreateRibbonGroup();
            this.editShapeGroup = this.Factory.CreateRibbonGroup();
            this.ResizeTextBox = this.Factory.CreateRibbonEditBox();
            this.etcGroup = this.Factory.CreateRibbonGroup();
            this.UpdateButton = this.Factory.CreateRibbonButton();
            this.ReplaceButton = this.Factory.CreateRibbonButton();
            this.ZebraButton = this.Factory.CreateRibbonButton();
            this.RemoveEmptyColButton = this.Factory.CreateRibbonButton();
            this.RemoveEmptyRowButton = this.Factory.CreateRibbonButton();
            this.AggregateButton = this.Factory.CreateRibbonButton();
            this.MergeFormatCondsButton = this.Factory.CreateRibbonButton();
            this.ShapeEditButton = this.Factory.CreateRibbonButton();
            this.ResizeButton = this.Factory.CreateRibbonButton();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.TestControlButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.updateGroup.SuspendLayout();
            this.editSheetGroup.SuspendLayout();
            this.editShapeGroup.SuspendLayout();
            this.etcGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.updateGroup);
            this.tab1.Groups.Add(this.editSheetGroup);
            this.tab1.Groups.Add(this.editShapeGroup);
            this.tab1.Groups.Add(this.etcGroup);
            this.tab1.Label = "SSC";
            this.tab1.Name = "tab1";
            // 
            // updateGroup
            // 
            this.updateGroup.Items.Add(this.UpdateButton);
            this.updateGroup.Label = "アップデート";
            this.updateGroup.Name = "updateGroup";
            this.updateGroup.Visible = false;
            // 
            // editSheetGroup
            // 
            this.editSheetGroup.Items.Add(this.ReplaceButton);
            this.editSheetGroup.Items.Add(this.ZebraButton);
            this.editSheetGroup.Items.Add(this.RemoveEmptyColButton);
            this.editSheetGroup.Items.Add(this.RemoveEmptyRowButton);
            this.editSheetGroup.Items.Add(this.AggregateButton);
            this.editSheetGroup.Items.Add(this.MergeFormatCondsButton);
            this.editSheetGroup.Label = "シート編集";
            this.editSheetGroup.Name = "editSheetGroup";
            // 
            // editShapeGroup
            // 
            this.editShapeGroup.Items.Add(this.ShapeEditButton);
            this.editShapeGroup.Items.Add(this.ResizeButton);
            this.editShapeGroup.Items.Add(this.ResizeTextBox);
            this.editShapeGroup.Label = "シェイプ編集";
            this.editShapeGroup.Name = "editShapeGroup";
            // 
            // ResizeTextBox
            // 
            this.ResizeTextBox.Label = "倍率";
            this.ResizeTextBox.MaxLength = 3;
            this.ResizeTextBox.Name = "ResizeTextBox";
            this.ResizeTextBox.SizeString = "000";
            this.ResizeTextBox.Text = null;
            this.ResizeTextBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ResizeTextBox_TextChanged);
            // 
            // etcGroup
            // 
            this.etcGroup.Items.Add(this.AboutButton);
            this.etcGroup.Items.Add(this.TestControlButton);
            this.etcGroup.Label = "etc";
            this.etcGroup.Name = "etcGroup";
            // 
            // UpdateButton
            // 
            this.UpdateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdateButton.Image = global::SscExcelAddIn.Properties.Resources.icons8_double_exclamation_mark_96;
            this.UpdateButton.Label = "更新が\nあります";
            this.UpdateButton.Name = "UpdateButton";
            this.UpdateButton.ShowImage = true;
            this.UpdateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateButton_Click);
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
            // RemoveEmptyColButton
            // 
            this.RemoveEmptyColButton.Image = global::SscExcelAddIn.Properties.Resources.icons8_delete_column_96;
            this.RemoveEmptyColButton.Label = "空列削除";
            this.RemoveEmptyColButton.Name = "RemoveEmptyColButton";
            this.RemoveEmptyColButton.ShowImage = true;
            this.RemoveEmptyColButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveEmptyColButton_Click);
            // 
            // RemoveEmptyRowButton
            // 
            this.RemoveEmptyRowButton.Image = global::SscExcelAddIn.Properties.Resources.icons8_delete_row_96;
            this.RemoveEmptyRowButton.Label = "空行削除";
            this.RemoveEmptyRowButton.Name = "RemoveEmptyRowButton";
            this.RemoveEmptyRowButton.ShowImage = true;
            this.RemoveEmptyRowButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveEmptyRowButton_Click);
            // 
            // AggregateButton
            // 
            this.AggregateButton.Image = global::SscExcelAddIn.Properties.Resources.icons8_table_96;
            this.AggregateButton.Label = "集計表";
            this.AggregateButton.Name = "AggregateButton";
            this.AggregateButton.ShowImage = true;
            this.AggregateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AggregateButton_Click);
            // 
            // MergeFormatCondsButton
            // 
            this.MergeFormatCondsButton.Image = global::SscExcelAddIn.Properties.Resources.icons8_compose_96;
            this.MergeFormatCondsButton.Label = "条件書式整理";
            this.MergeFormatCondsButton.Name = "MergeFormatCondsButton";
            this.MergeFormatCondsButton.ShowImage = true;
            this.MergeFormatCondsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MergeFormatCondsButton_Click);
            // 
            // ShapeEditButton
            // 
            this.ShapeEditButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShapeEditButton.Image = global::SscExcelAddIn.Properties.Resources.shapes_icon_128261;
            this.ShapeEditButton.Label = "シェイプ\n文字列";
            this.ShapeEditButton.Name = "ShapeEditButton";
            this.ShapeEditButton.ShowImage = true;
            this.ShapeEditButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShapeEditButton_Click);
            // 
            // ResizeButton
            // 
            this.ResizeButton.Image = global::SscExcelAddIn.Properties.Resources.resize_full_icon_178778;
            this.ResizeButton.Label = "リサイズ";
            this.ResizeButton.Name = "ResizeButton";
            this.ResizeButton.ShowImage = true;
            this.ResizeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ResizeButton_Click);
            // 
            // AboutButton
            // 
            this.AboutButton.Label = "About";
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
            this.updateGroup.ResumeLayout(false);
            this.updateGroup.PerformLayout();
            this.editSheetGroup.ResumeLayout(false);
            this.editSheetGroup.PerformLayout();
            this.editShapeGroup.ResumeLayout(false);
            this.editShapeGroup.PerformLayout();
            this.etcGroup.ResumeLayout(false);
            this.etcGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup editSheetGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReplaceButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ZebraButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup etcGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShapeEditButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup editShapeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ResizeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ResizeTextBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup updateGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveEmptyColButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveEmptyRowButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AggregateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MergeFormatCondsButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
