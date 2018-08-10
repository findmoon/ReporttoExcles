namespace Exicel转换1
{
    partial class QuotationLackItem
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.lackItemTxt = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.DarkRed;
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(329, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "生成的报价单缺少以下内容，请自行补充或联系软件制作者：";
            // 
            // lackItemTxt
            // 
            this.lackItemTxt.Location = new System.Drawing.Point(14, 58);
            this.lackItemTxt.Multiline = true;
            this.lackItemTxt.Name = "lackItemTxt";
            this.lackItemTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.lackItemTxt.Size = new System.Drawing.Size(339, 366);
            this.lackItemTxt.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(374, 36);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(64, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "确定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // QuotationLackItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(450, 446);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lackItemTxt);
            this.Controls.Add(this.label1);
            this.Name = "QuotationLackItem";
            this.Text = "QuotationLackItem";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox lackItemTxt;
        private System.Windows.Forms.Button button1;
    }
}