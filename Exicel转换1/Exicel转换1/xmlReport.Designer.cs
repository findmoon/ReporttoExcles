﻿namespace Exicel转换1
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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("最近转换的Job Report");
            this.label1 = new System.Windows.Forms.Label();
            this.OpenXmlFolder_Btn = new System.Windows.Forms.Button();
            this.GenExcleReport_Btn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.HasOpenFolderLbel = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.latestEvaluationReports_treeView = new System.Windows.Forms.TreeView();
            this.evaluation_TabControl = new System.Windows.Forms.TabControl();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(98, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "请选择flexa导出的xml文件夹：";
            // 
            // OpenXmlFolder_Btn
            // 
            this.OpenXmlFolder_Btn.Location = new System.Drawing.Point(310, 22);
            this.OpenXmlFolder_Btn.Name = "OpenXmlFolder_Btn";
            this.OpenXmlFolder_Btn.Size = new System.Drawing.Size(108, 36);
            this.OpenXmlFolder_Btn.TabIndex = 1;
            this.OpenXmlFolder_Btn.Text = "打开";
            this.OpenXmlFolder_Btn.UseVisualStyleBackColor = true;
            this.OpenXmlFolder_Btn.Click += new System.EventHandler(this.OpenXmlFolder_Btn_Click);
            // 
            // GenExcleReport_Btn
            // 
            this.GenExcleReport_Btn.Location = new System.Drawing.Point(514, 23);
            this.GenExcleReport_Btn.Name = "GenExcleReport_Btn";
            this.GenExcleReport_Btn.Size = new System.Drawing.Size(108, 36);
            this.GenExcleReport_Btn.TabIndex = 2;
            this.GenExcleReport_Btn.Text = "导出";
            this.GenExcleReport_Btn.UseVisualStyleBackColor = true;
            this.GenExcleReport_Btn.Click += new System.EventHandler(this.GenExcleReport_Btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(98, 57);
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
            // splitContainer2
            // 
            this.splitContainer2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer2.Location = new System.Drawing.Point(13, 96);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.latestEvaluationReports_treeView);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.evaluation_TabControl);
            this.splitContainer2.Size = new System.Drawing.Size(1084, 600);
            this.splitContainer2.SplitterDistance = 176;
            this.splitContainer2.SplitterWidth = 1;
            this.splitContainer2.TabIndex = 6;
            // 
            // latestEvaluationReports_treeView
            // 
            this.latestEvaluationReports_treeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.latestEvaluationReports_treeView.Location = new System.Drawing.Point(0, 0);
            this.latestEvaluationReports_treeView.Name = "latestEvaluationReports_treeView";
            treeNode1.Name = "root";
            treeNode1.Text = "最近转换的Job Report";
            this.latestEvaluationReports_treeView.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.latestEvaluationReports_treeView.Size = new System.Drawing.Size(176, 600);
            this.latestEvaluationReports_treeView.TabIndex = 2;
            this.latestEvaluationReports_treeView.DoubleClick += new System.EventHandler(this.latestEvaluationReports_treeView_DoubleClick);
            // 
            // evaluation_TabControl
            // 
            this.evaluation_TabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.evaluation_TabControl.Location = new System.Drawing.Point(2, 0);
            this.evaluation_TabControl.Margin = new System.Windows.Forms.Padding(0);
            this.evaluation_TabControl.Name = "evaluation_TabControl";
            this.evaluation_TabControl.Padding = new System.Drawing.Point(1, 1);
            this.evaluation_TabControl.SelectedIndex = 0;
            this.evaluation_TabControl.Size = new System.Drawing.Size(905, 890);
            this.evaluation_TabControl.TabIndex = 4;
            // 
            // XmlReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.splitContainer2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.GenExcleReport_Btn);
            this.Controls.Add(this.OpenXmlFolder_Btn);
            this.Controls.Add(this.HasOpenFolderLbel);
            this.Controls.Add(this.label1);
            this.Name = "XmlReport";
            this.Size = new System.Drawing.Size(1100, 700);
            this.Load += new System.EventHandler(this.XmlReport_Load);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
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
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.TreeView latestEvaluationReports_treeView;
        private System.Windows.Forms.TabControl evaluation_TabControl;
    }
}
