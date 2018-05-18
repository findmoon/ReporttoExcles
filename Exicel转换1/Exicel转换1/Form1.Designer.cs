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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.GpbWindow = new System.Windows.Forms.GroupBox();
            this.textLog = new System.Windows.Forms.TextBox();
            this.FlexCVWindow = new System.Windows.Forms.Button();
            this.XmlWindow = new System.Windows.Forms.Button();
            this.LinCfgWindow = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // GpbWindow
            // 
            this.GpbWindow.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GpbWindow.AutoSize = true;
            this.GpbWindow.Location = new System.Drawing.Point(12, 78);
            this.GpbWindow.Name = "GpbWindow";
            this.GpbWindow.Size = new System.Drawing.Size(460, 275);
            this.GpbWindow.TabIndex = 3;
            this.GpbWindow.TabStop = false;
            this.GpbWindow.Enter += new System.EventHandler(this.GpbWindow_Enter);
            // 
            // textLog
            // 
            this.textLog.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textLog.Location = new System.Drawing.Point(3, 409);
            this.textLog.Name = "textLog";
            this.textLog.Size = new System.Drawing.Size(460, 21);
            this.textLog.TabIndex = 0;
            // 
            // FlexCVWindow
            // 
            this.FlexCVWindow.Location = new System.Drawing.Point(108, 12);
            this.FlexCVWindow.Name = "FlexCVWindow";
            this.FlexCVWindow.Size = new System.Drawing.Size(82, 23);
            this.FlexCVWindow.TabIndex = 1;
            this.FlexCVWindow.Text = "导入FlexCV report";
            this.FlexCVWindow.UseVisualStyleBackColor = true;
            this.FlexCVWindow.Click += new System.EventHandler(this.FlexCVWindow_Click);
            // 
            // XmlWindow
            // 
            this.XmlWindow.Location = new System.Drawing.Point(12, 12);
            this.XmlWindow.Name = "XmlWindow";
            this.XmlWindow.Size = new System.Drawing.Size(90, 23);
            this.XmlWindow.TabIndex = 2;
            this.XmlWindow.Text = "导入xml";
            this.XmlWindow.UseVisualStyleBackColor = true;
            this.XmlWindow.Click += new System.EventHandler(this.XmlWindow_Click);
            // 
            // LinCfgWindow
            // 
            this.LinCfgWindow.Location = new System.Drawing.Point(196, 12);
            this.LinCfgWindow.Name = "LinCfgWindow";
            this.LinCfgWindow.Size = new System.Drawing.Size(112, 23);
            this.LinCfgWindow.TabIndex = 4;
            this.LinCfgWindow.Text = "导入LineConfg";
            this.LinCfgWindow.UseVisualStyleBackColor = true;
            this.LinCfgWindow.Click += new System.EventHandler(this.LinCfgWindow_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(475, 423);
            this.Controls.Add(this.textLog);
            this.Controls.Add(this.LinCfgWindow);
            this.Controls.Add(this.FlexCVWindow);
            this.Controls.Add(this.XmlWindow);
            this.Controls.Add(this.GpbWindow);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "转换导出job report";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button LinCfgWindow;
        private System.Windows.Forms.Button FlexCVWindow;
        private System.Windows.Forms.Button XmlWindow;
        private System.Windows.Forms.GroupBox GpbWindow;
        private System.Windows.Forms.TextBox textLog;
    }
}

