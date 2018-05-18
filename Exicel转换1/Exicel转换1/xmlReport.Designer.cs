namespace Exicel转换1
{
    partial class XmlReport
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
            this.label1 = new System.Windows.Forms.Label();
            this.OpenXmlFolder_Btn = new System.Windows.Forms.Button();
            this.GenExcleReport_Btn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.HasOpenFolderLbel = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "请选择flexa导出的xml文件夹：";
            // 
            // OpenXmlFolder_Btn
            // 
            this.OpenXmlFolder_Btn.Location = new System.Drawing.Point(58, 47);
            this.OpenXmlFolder_Btn.Name = "OpenXmlFolder_Btn";
            this.OpenXmlFolder_Btn.Size = new System.Drawing.Size(122, 51);
            this.OpenXmlFolder_Btn.TabIndex = 1;
            this.OpenXmlFolder_Btn.Text = "打开";
            this.OpenXmlFolder_Btn.UseVisualStyleBackColor = true;
            this.OpenXmlFolder_Btn.Click += new System.EventHandler(this.OpenXmlFolder_Btn_Click);
            // 
            // GenExcleReport_Btn
            // 
            this.GenExcleReport_Btn.Location = new System.Drawing.Point(319, 299);
            this.GenExcleReport_Btn.Name = "GenExcleReport_Btn";
            this.GenExcleReport_Btn.Size = new System.Drawing.Size(122, 49);
            this.GenExcleReport_Btn.TabIndex = 2;
            this.GenExcleReport_Btn.Text = "导出";
            this.GenExcleReport_Btn.UseVisualStyleBackColor = true;
            this.GenExcleReport_Btn.Click += new System.EventHandler(this.GenExcleReport_Btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 128);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "已打开的文件夹：";
            // 
            // HasOpenFolderLbel
            // 
            this.HasOpenFolderLbel.AutoSize = true;
            this.HasOpenFolderLbel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.HasOpenFolderLbel.ForeColor = System.Drawing.SystemColors.InfoText;
            this.HasOpenFolderLbel.Location = new System.Drawing.Point(19, 153);
            this.HasOpenFolderLbel.Name = "HasOpenFolderLbel";
            this.HasOpenFolderLbel.Size = new System.Drawing.Size(0, 12);
            this.HasOpenFolderLbel.TabIndex = 4;
            // 
            // XmlReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.HasOpenFolderLbel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.GenExcleReport_Btn);
            this.Controls.Add(this.OpenXmlFolder_Btn);
            this.Controls.Add(this.label1);
            this.Name = "XmlReport";
            this.Size = new System.Drawing.Size(467, 372);
            this.Load += new System.EventHandler(this.XmlReport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button OpenXmlFolder_Btn;
        private System.Windows.Forms.Button GenExcleReport_Btn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label HasOpenFolderLbel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}
