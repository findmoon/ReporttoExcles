namespace Exicel转换1
{
    partial class FlexCVReport
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ExcelToDTbut = new System.Windows.Forms.Button();
            this.GenExlbut = new System.Windows.Forms.Button();
            this.openExcelFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.SuspendLayout();
            // 
            // ExcelToDTbut
            // 
            this.ExcelToDTbut.Location = new System.Drawing.Point(168, 69);
            this.ExcelToDTbut.Name = "ExcelToDTbut";
            this.ExcelToDTbut.Size = new System.Drawing.Size(142, 58);
            this.ExcelToDTbut.TabIndex = 4;
            this.ExcelToDTbut.Text = "打开";
            this.ExcelToDTbut.UseVisualStyleBackColor = true;
            this.ExcelToDTbut.Click += new System.EventHandler(this.ExcelToDTbut_Click);
            // 
            // GenExlbut
            // 
            this.GenExlbut.Location = new System.Drawing.Point(168, 192);
            this.GenExlbut.Name = "GenExlbut";
            this.GenExlbut.Size = new System.Drawing.Size(142, 56);
            this.GenExlbut.TabIndex = 3;
            this.GenExlbut.Text = "生成Excel";
            this.GenExlbut.UseVisualStyleBackColor = true;
            this.GenExlbut.Click += new System.EventHandler(this.GenExlbut_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // FlexCVReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.ExcelToDTbut);
            this.Controls.Add(this.GenExlbut);
            this.Name = "FlexCVReport";
            this.Size = new System.Drawing.Size(518, 359);
            this.Load += new System.EventHandler(this.importFlexCVCReport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ExcelToDTbut;
        private System.Windows.Forms.Button GenExlbut;
        private System.Windows.Forms.OpenFileDialog openExcelFileDialog;
        private System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
    }
}
