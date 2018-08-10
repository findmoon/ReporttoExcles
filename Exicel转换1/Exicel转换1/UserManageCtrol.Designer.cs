namespace Exicel转换1
{
    partial class UserManageCtrol
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
            this.UserAdminTab = new System.Windows.Forms.TabControl();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // UserAdminTab
            // 
            this.UserAdminTab.Location = new System.Drawing.Point(56, 47);
            this.UserAdminTab.Margin = new System.Windows.Forms.Padding(12, 3, 3, 3);
            this.UserAdminTab.Name = "UserAdminTab";
            this.UserAdminTab.SelectedIndex = 0;
            this.UserAdminTab.Size = new System.Drawing.Size(807, 359);
            this.UserAdminTab.TabIndex = 0;
            this.UserAdminTab.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.UserAdminTab_DrawItem);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(230, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(400, 44);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户管理";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // UserManageCtrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.UserAdminTab);
            this.Name = "UserManageCtrol";
            this.Size = new System.Drawing.Size(900, 409);
            this.Load += new System.EventHandler(this.UserManageCtrol_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl UserAdminTab;
        private System.Windows.Forms.Label label1;
    }
}
