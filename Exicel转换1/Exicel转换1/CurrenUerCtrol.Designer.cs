namespace Exicel转换1
{
    partial class CurrenUerCtrol
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
            this.userNameTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.oldpwdTxt = new System.Windows.Forms.TextBox();
            this.newpwdTxt = new System.Windows.Forms.TextBox();
            this.confirmNewpwdTxt = new System.Windows.Forms.TextBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.SaveBtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(73, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户名：";
            // 
            // userNameTxt
            // 
            this.userNameTxt.Location = new System.Drawing.Point(157, 29);
            this.userNameTxt.Name = "userNameTxt";
            this.userNameTxt.Size = new System.Drawing.Size(100, 21);
            this.userNameTxt.TabIndex = 1;
            this.userNameTxt.TextChanged += new System.EventHandler(this.userNameTxt_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "旧密码：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "新密码：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 123);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 12);
            this.label4.TabIndex = 4;
            this.label4.Text = "确认新密码：";
            // 
            // oldpwdTxt
            // 
            this.oldpwdTxt.Location = new System.Drawing.Point(100, 40);
            this.oldpwdTxt.Name = "oldpwdTxt";
            this.oldpwdTxt.Size = new System.Drawing.Size(100, 21);
            this.oldpwdTxt.TabIndex = 5;
            // 
            // newpwdTxt
            // 
            this.newpwdTxt.Location = new System.Drawing.Point(100, 80);
            this.newpwdTxt.Name = "newpwdTxt";
            this.newpwdTxt.Size = new System.Drawing.Size(100, 21);
            this.newpwdTxt.TabIndex = 6;
            // 
            // confirmNewpwdTxt
            // 
            this.confirmNewpwdTxt.Location = new System.Drawing.Point(100, 114);
            this.confirmNewpwdTxt.Name = "confirmNewpwdTxt";
            this.confirmNewpwdTxt.Size = new System.Drawing.Size(100, 21);
            this.confirmNewpwdTxt.TabIndex = 7;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(264, 276);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(29, 12);
            this.linkLabel1.TabIndex = 8;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "编辑";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // SaveBtn
            // 
            this.SaveBtn.Location = new System.Drawing.Point(96, 265);
            this.SaveBtn.Name = "SaveBtn";
            this.SaveBtn.Size = new System.Drawing.Size(75, 23);
            this.SaveBtn.TabIndex = 9;
            this.SaveBtn.Text = "保存";
            this.SaveBtn.UseVisualStyleBackColor = true;
            this.SaveBtn.Click += new System.EventHandler(this.SaveBtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.confirmNewpwdTxt);
            this.groupBox1.Controls.Add(this.oldpwdTxt);
            this.groupBox1.Controls.Add(this.newpwdTxt);
            this.groupBox1.Location = new System.Drawing.Point(56, 71);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(215, 148);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "修改密码";
            // 
            // CurrenUerCtrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.SaveBtn);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.userNameTxt);
            this.Controls.Add(this.label1);
            this.Name = "CurrenUerCtrol";
            this.Size = new System.Drawing.Size(381, 327);
            this.Load += new System.EventHandler(this.CurrenUerCtrol_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox userNameTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox oldpwdTxt;
        private System.Windows.Forms.TextBox newpwdTxt;
        private System.Windows.Forms.TextBox confirmNewpwdTxt;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button SaveBtn;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}
