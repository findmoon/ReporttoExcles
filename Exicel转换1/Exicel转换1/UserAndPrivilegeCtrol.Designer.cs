namespace Exicel转换1
{
    partial class UserAndPrivilegeCtrol
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.userNameTxt = new System.Windows.Forms.TextBox();
            this.pwdTxt = new System.Windows.Forms.TextBox();
            this.confirmPwdTxt = new System.Windows.Forms.TextBox();
            this.PrivilegeCombo = new System.Windows.Forms.ComboBox();
            this.AddUserBtn = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.AllUserCombo = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.modefyUserNameTxt = new System.Windows.Forms.TextBox();
            this.modefyPwdTxt = new System.Windows.Forms.TextBox();
            this.modefyPrivilegeCombo = new System.Windows.Forms.ComboBox();
            this.ModefyBtn = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ModefyBtn);
            this.groupBox1.Controls.Add(this.modefyPrivilegeCombo);
            this.groupBox1.Controls.Add(this.modefyPwdTxt);
            this.groupBox1.Controls.Add(this.modefyUserNameTxt);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.AllUserCombo);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(24, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(495, 157);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "修改用户";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.AddUserBtn);
            this.groupBox2.Controls.Add(this.PrivilegeCombo);
            this.groupBox2.Controls.Add(this.confirmPwdTxt);
            this.groupBox2.Controls.Add(this.pwdTxt);
            this.groupBox2.Controls.Add(this.userNameTxt);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(24, 191);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(495, 164);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "添加用户";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户名：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "密码：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "确认密码：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 131);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "权限：";
            // 
            // userNameTxt
            // 
            this.userNameTxt.Location = new System.Drawing.Point(100, 27);
            this.userNameTxt.Name = "userNameTxt";
            this.userNameTxt.Size = new System.Drawing.Size(118, 21);
            this.userNameTxt.TabIndex = 4;
            // 
            // pwdTxt
            // 
            this.pwdTxt.Location = new System.Drawing.Point(100, 64);
            this.pwdTxt.Name = "pwdTxt";
            this.pwdTxt.PasswordChar = '*';
            this.pwdTxt.Size = new System.Drawing.Size(118, 21);
            this.pwdTxt.TabIndex = 5;
            // 
            // confirmPwdTxt
            // 
            this.confirmPwdTxt.Location = new System.Drawing.Point(100, 97);
            this.confirmPwdTxt.Name = "confirmPwdTxt";
            this.confirmPwdTxt.PasswordChar = '*';
            this.confirmPwdTxt.Size = new System.Drawing.Size(118, 21);
            this.confirmPwdTxt.TabIndex = 6;
            // 
            // PrivilegeCombo
            // 
            this.PrivilegeCombo.FormattingEnabled = true;
            this.PrivilegeCombo.Location = new System.Drawing.Point(100, 131);
            this.PrivilegeCombo.Name = "PrivilegeCombo";
            this.PrivilegeCombo.Size = new System.Drawing.Size(118, 20);
            this.PrivilegeCombo.TabIndex = 7;
            // 
            // AddUserBtn
            // 
            this.AddUserBtn.Location = new System.Drawing.Point(303, 100);
            this.AddUserBtn.Name = "AddUserBtn";
            this.AddUserBtn.Size = new System.Drawing.Size(69, 30);
            this.AddUserBtn.TabIndex = 8;
            this.AddUserBtn.Text = "添加";
            this.AddUserBtn.UseVisualStyleBackColor = true;
            this.AddUserBtn.Click += new System.EventHandler(this.AddUserBtn_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 14);
            this.label5.TabIndex = 0;
            this.label5.Text = "选择用户：";
            // 
            // AllUserCombo
            // 
            this.AllUserCombo.FormattingEnabled = true;
            this.AllUserCombo.Location = new System.Drawing.Point(82, 20);
            this.AllUserCombo.Name = "AllUserCombo";
            this.AllUserCombo.Size = new System.Drawing.Size(121, 22);
            this.AllUserCombo.TabIndex = 1;
            this.AllUserCombo.SelectedIndexChanged += new System.EventHandler(this.AllUserCombo_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(79, 61);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(91, 14);
            this.label6.TabIndex = 2;
            this.label6.Text = "修改用户名：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(79, 128);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 14);
            this.label7.TabIndex = 3;
            this.label7.Text = "修改权限：";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(79, 96);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 14);
            this.label8.TabIndex = 4;
            this.label8.Text = "修改密码：";
            // 
            // modefyUserNameTxt
            // 
            this.modefyUserNameTxt.Location = new System.Drawing.Point(187, 58);
            this.modefyUserNameTxt.Name = "modefyUserNameTxt";
            this.modefyUserNameTxt.Size = new System.Drawing.Size(120, 23);
            this.modefyUserNameTxt.TabIndex = 5;
            // 
            // modefyPwdTxt
            // 
            this.modefyPwdTxt.Location = new System.Drawing.Point(187, 90);
            this.modefyPwdTxt.Name = "modefyPwdTxt";
            this.modefyPwdTxt.PasswordChar = '*';
            this.modefyPwdTxt.Size = new System.Drawing.Size(120, 23);
            this.modefyPwdTxt.TabIndex = 6;
            // 
            // modefyPrivilegeCombo
            // 
            this.modefyPrivilegeCombo.FormattingEnabled = true;
            this.modefyPrivilegeCombo.Location = new System.Drawing.Point(187, 125);
            this.modefyPrivilegeCombo.Name = "modefyPrivilegeCombo";
            this.modefyPrivilegeCombo.Size = new System.Drawing.Size(121, 22);
            this.modefyPrivilegeCombo.TabIndex = 8;
            // 
            // ModefyBtn
            // 
            this.ModefyBtn.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ModefyBtn.Location = new System.Drawing.Point(406, 128);
            this.ModefyBtn.Name = "ModefyBtn";
            this.ModefyBtn.Size = new System.Drawing.Size(68, 28);
            this.ModefyBtn.TabIndex = 9;
            this.ModefyBtn.Text = "确认修改";
            this.ModefyBtn.UseVisualStyleBackColor = true;
            this.ModefyBtn.Click += new System.EventHandler(this.ModefyBtn_Click);
            // 
            // UserAndPrivilegeCtrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "UserAndPrivilegeCtrol";
            this.Size = new System.Drawing.Size(548, 370);
            this.Load += new System.EventHandler(this.UserAndPrivilegeCtrol_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox PrivilegeCombo;
        private System.Windows.Forms.TextBox confirmPwdTxt;
        private System.Windows.Forms.TextBox pwdTxt;
        private System.Windows.Forms.TextBox userNameTxt;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox modefyUserNameTxt;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox AllUserCombo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button AddUserBtn;
        private System.Windows.Forms.Button ModefyBtn;
        private System.Windows.Forms.ComboBox modefyPrivilegeCombo;
        private System.Windows.Forms.TextBox modefyPwdTxt;
    }
}
