namespace Exicel转换1
{
    partial class Form1
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.GenExlbut = new System.Windows.Forms.Button();
            this.ExcelToDTbut = new System.Windows.Forms.Button();
            this.openExcelFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // GenExlbut
            // 
            this.GenExlbut.Location = new System.Drawing.Point(71, 189);
            this.GenExlbut.Name = "GenExlbut";
            this.GenExlbut.Size = new System.Drawing.Size(142, 56);
            this.GenExlbut.TabIndex = 1;
            this.GenExlbut.Text = "生成Excel";
            this.GenExlbut.UseVisualStyleBackColor = true;
            this.GenExlbut.Click += new System.EventHandler(this.GenExlbut_Click);
            // 
            // ExcelToDTbut
            // 
            this.ExcelToDTbut.Location = new System.Drawing.Point(71, 37);
            this.ExcelToDTbut.Name = "ExcelToDTbut";
            this.ExcelToDTbut.Size = new System.Drawing.Size(142, 58);
            this.ExcelToDTbut.TabIndex = 2;
            this.ExcelToDTbut.Text = "打开";
            this.ExcelToDTbut.UseVisualStyleBackColor = true;
            this.ExcelToDTbut.Click += new System.EventHandler(this.ExcelToDTbut_Click);
            // 
            // bindingSource1
            // 
            this.bindingSource1.CurrentChanged += new System.EventHandler(this.bindingSource1_CurrentChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(288, 297);
            this.Controls.Add(this.ExcelToDTbut);
            this.Controls.Add(this.GenExlbut);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "转换导出job report";
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button GenExlbut;
        private System.Windows.Forms.Button ExcelToDTbut;
        private System.Windows.Forms.OpenFileDialog openExcelFileDialog;
        private System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
        private System.Windows.Forms.BindingSource bindingSource1;
    }
}

