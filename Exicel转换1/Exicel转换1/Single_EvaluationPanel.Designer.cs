namespace Exicel转换1
{
    partial class Single_EvaluationPanelControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Single_EvaluationPanelControl));
            this.Single_Evaluation = new System.Windows.Forms.Panel();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.feederNozzleDGV = new System.Windows.Forms.DataGridView();
            this.summaryDGV = new System.Windows.Forms.DataGridView();
            this.layoutDGV = new System.Windows.Forms.DataGridView();
            this.Single_Evaluation.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.feederNozzleDGV)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.summaryDGV)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // Single_Evaluation
            // 
            this.Single_Evaluation.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Single_Evaluation.AutoScroll = true;
            this.Single_Evaluation.Controls.Add(this.toolStrip1);
            this.Single_Evaluation.Controls.Add(this.feederNozzleDGV);
            this.Single_Evaluation.Controls.Add(this.summaryDGV);
            this.Single_Evaluation.Controls.Add(this.layoutDGV);
            this.Single_Evaluation.Location = new System.Drawing.Point(0, 0);
            this.Single_Evaluation.Name = "Single_Evaluation";
            this.Single_Evaluation.Size = new System.Drawing.Size(845, 319);
            this.Single_Evaluation.TabIndex = 14;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.toolStripSeparator1,
            this.toolStripLabel2,
            this.toolStripButton2});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(828, 25);
            this.toolStrip1.TabIndex = 19;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(84, 22);
            this.toolStripLabel1.Text = "修改Base组合";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(123, 22);
            this.toolStripLabel2.Text = "修改Head和理论CPH";
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton2.Text = "toolStripButton2";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // feederNozzleDGV
            // 
            this.feederNozzleDGV.AllowUserToAddRows = false;
            this.feederNozzleDGV.AllowUserToDeleteRows = false;
            this.feederNozzleDGV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.feederNozzleDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.feederNozzleDGV.Location = new System.Drawing.Point(498, 173);
            this.feederNozzleDGV.Name = "feederNozzleDGV";
            this.feederNozzleDGV.RowTemplate.Height = 23;
            this.feederNozzleDGV.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.feederNozzleDGV.Size = new System.Drawing.Size(330, 181);
            this.feederNozzleDGV.TabIndex = 18;
            // 
            // summaryDGV
            // 
            this.summaryDGV.AllowUserToAddRows = false;
            this.summaryDGV.AllowUserToDeleteRows = false;
            this.summaryDGV.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.summaryDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.summaryDGV.Location = new System.Drawing.Point(4, 38);
            this.summaryDGV.Name = "summaryDGV";
            this.summaryDGV.RowTemplate.Height = 23;
            this.summaryDGV.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.summaryDGV.Size = new System.Drawing.Size(824, 112);
            this.summaryDGV.TabIndex = 17;
            // 
            // layoutDGV
            // 
            this.layoutDGV.AllowUserToAddRows = false;
            this.layoutDGV.AllowUserToDeleteRows = false;
            this.layoutDGV.AllowUserToOrderColumns = true;
            this.layoutDGV.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.layoutDGV.Location = new System.Drawing.Point(0, 173);
            this.layoutDGV.Name = "layoutDGV";
            this.layoutDGV.RowTemplate.Height = 23;
            this.layoutDGV.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.layoutDGV.Size = new System.Drawing.Size(499, 181);
            this.layoutDGV.TabIndex = 16;
            // 
            // Single_EvaluationPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.Single_Evaluation);
            this.Name = "Single_EvaluationPanelControl";
            this.Size = new System.Drawing.Size(845, 322);
            this.Single_Evaluation.ResumeLayout(false);
            this.Single_Evaluation.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.feederNozzleDGV)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.summaryDGV)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel Single_Evaluation;
        private System.Windows.Forms.DataGridView layoutDGV;
        private System.Windows.Forms.DataGridView feederNozzleDGV;
        private System.Windows.Forms.DataGridView summaryDGV;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
    }
}
