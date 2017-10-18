namespace Excel.AddIn.Shape2Image
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// デザイナー変数が必要です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
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
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.TabAddInsShape = this.Factory.CreateRibbonGroup();
            this.SaveAsPicture = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.TabAddInsShape.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.TabAddInsShape);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // TabAddInsShape
            // 
            this.TabAddInsShape.Items.Add(this.SaveAsPicture);
            this.TabAddInsShape.Label = "図";
            this.TabAddInsShape.Name = "TabAddInsShape";
            // 
            // SaveAsPicture
            // 
            this.SaveAsPicture.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SaveAsPicture.Label = "名前を付けて保存";
            this.SaveAsPicture.Name = "SaveAsPicture";
            this.SaveAsPicture.OfficeImageId = "ImagerScan";
            this.SaveAsPicture.ShowImage = true;
            this.SaveAsPicture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsPicture_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.TabAddInsShape.ResumeLayout(false);
            this.TabAddInsShape.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TabAddInsShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SaveAsPicture;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
