namespace Exicel转换1
{
    partial class XmlExcelRtfConvertMain_Form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(XmlExcelRtfConvertMain_Form));
            this.GpbWindow = new System.Windows.Forms.GroupBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.XmlWindowToolBtn = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.FlexCVWindowToolBtn = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.LinCfgWindowToolBtn = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.UserAdmin = new System.Windows.Forms.ToolStripMenuItem();
            this.Quit = new System.Windows.Forms.ToolStripMenuItem();
            this.DetialLogBtn = new System.Windows.Forms.ToolStripButton();
            this.LogTxt = new System.Windows.Forms.TextBox();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // GpbWindow
            // 
            this.GpbWindow.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GpbWindow.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.GpbWindow.Location = new System.Drawing.Point(5, 35);
            this.GpbWindow.Margin = new System.Windows.Forms.Padding(3, 10, 3, 3);
            this.GpbWindow.MinimumSize = new System.Drawing.Size(800, 400);
            this.GpbWindow.Name = "GpbWindow";
            this.GpbWindow.Padding = new System.Windows.Forms.Padding(3, 10, 3, 3);
            this.GpbWindow.Size = new System.Drawing.Size(1109, 400);
            this.GpbWindow.TabIndex = 3;
            this.GpbWindow.TabStop = false;
            this.GpbWindow.Enter += new System.EventHandler(this.GpbWindow_Enter);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.XmlWindowToolBtn,
            this.toolStripSeparator1,
            this.FlexCVWindowToolBtn,
            this.toolStripSeparator2,
            this.LinCfgWindowToolBtn,
            this.toolStripButton1,
            this.DetialLogBtn});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1114, 25);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // XmlWindowToolBtn
            // 
            this.XmlWindowToolBtn.BackColor = System.Drawing.SystemColors.ControlLight;
            this.XmlWindowToolBtn.Image = ((System.Drawing.Image)(resources.GetObject("XmlWindowToolBtn.Image")));
            this.XmlWindowToolBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.XmlWindowToolBtn.Name = "XmlWindowToolBtn";
            this.XmlWindowToolBtn.Size = new System.Drawing.Size(149, 22);
            this.XmlWindowToolBtn.Text = "导入Flexa xml Report";
            this.XmlWindowToolBtn.Click += new System.EventHandler(this.XmlWindowToolBtn_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // FlexCVWindowToolBtn
            // 
            this.FlexCVWindowToolBtn.BackColor = System.Drawing.SystemColors.ControlLight;
            this.FlexCVWindowToolBtn.Image = ((System.Drawing.Image)(resources.GetObject("FlexCVWindowToolBtn.Image")));
            this.FlexCVWindowToolBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.FlexCVWindowToolBtn.Name = "FlexCVWindowToolBtn";
            this.FlexCVWindowToolBtn.Size = new System.Drawing.Size(134, 22);
            this.FlexCVWindowToolBtn.Text = "导入FlexCV Report";
            this.FlexCVWindowToolBtn.Click += new System.EventHandler(this.FlexCVWindowToolBtn_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // LinCfgWindowToolBtn
            // 
            this.LinCfgWindowToolBtn.BackColor = System.Drawing.SystemColors.ControlLight;
            this.LinCfgWindowToolBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.LinCfgWindowToolBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.LinCfgWindowToolBtn.Name = "LinCfgWindowToolBtn";
            this.LinCfgWindowToolBtn.Size = new System.Drawing.Size(120, 22);
            this.LinCfgWindowToolBtn.Text = "LineConfg和Layout";
            this.LinCfgWindowToolBtn.Click += new System.EventHandler(this.LinCfgWindowToolBtn_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.UserAdmin,
            this.Quit});
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(69, 22);
            this.toolStripButton1.Text = "账户管理";
            // 
            // UserAdmin
            // 
            this.UserAdmin.Image = ((System.Drawing.Image)(resources.GetObject("UserAdmin.Image")));
            this.UserAdmin.Name = "UserAdmin";
            this.UserAdmin.Size = new System.Drawing.Size(124, 22);
            this.UserAdmin.Text = "查看管理";
            this.UserAdmin.Click += new System.EventHandler(this.UserAdmin_Click);
            // 
            // Quit
            // 
            this.Quit.Image = ((System.Drawing.Image)(resources.GetObject("Quit.Image")));
            this.Quit.Name = "Quit";
            this.Quit.Size = new System.Drawing.Size(124, 22);
            this.Quit.Text = "退出(&X)";
            // 
            // DetialLogBtn
            // 
            this.DetialLogBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.DetialLogBtn.Image = ((System.Drawing.Image)(resources.GetObject("DetialLogBtn.Image")));
            this.DetialLogBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.DetialLogBtn.Name = "DetialLogBtn";
            this.DetialLogBtn.Size = new System.Drawing.Size(60, 22);
            this.DetialLogBtn.Text = "详细记录";
            this.DetialLogBtn.Click += new System.EventHandler(this.DetialLogBtn_Click);
            // 
            // LogTxt
            // 
            this.LogTxt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LogTxt.Location = new System.Drawing.Point(5, 461);
            this.LogTxt.Multiline = true;
            this.LogTxt.Name = "LogTxt";
            this.LogTxt.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.LogTxt.Size = new System.Drawing.Size(1109, 21);
            this.LogTxt.TabIndex = 0;
            // 
            // XmlExcelRtfConvertMain_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1114, 554);
            this.Controls.Add(this.LogTxt);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.GpbWindow);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "XmlExcelRtfConvertMain_Form";
            this.Text = "转换导出job report";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox GpbWindow;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton XmlWindowToolBtn;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton LinCfgWindowToolBtn;
        private System.Windows.Forms.ToolStripDropDownButton toolStripButton1;
        private System.Windows.Forms.ToolStripMenuItem UserAdmin;
        private System.Windows.Forms.ToolStripMenuItem Quit;
        private System.Windows.Forms.ToolStripButton DetialLogBtn;
        private System.Windows.Forms.TextBox LogTxt;
        private System.Windows.Forms.ToolStripButton FlexCVWindowToolBtn;
    }
}

