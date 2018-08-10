namespace Exicel转换1
{
    partial class CPPCalculate
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
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cppLabel = new System.Windows.Forms.Label();
            this.hourComb = new System.Windows.Forms.ComboBox();
            this.dayComb = new System.Windows.Forms.ComboBox();
            this.monthComb = new System.Windows.Forms.ComboBox();
            this.yearComb = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "每天：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "每月：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "每年：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(28, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "年数：";
            // 
            // cppLabel
            // 
            this.cppLabel.AutoSize = true;
            this.cppLabel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cppLabel.Location = new System.Drawing.Point(220, 52);
            this.cppLabel.Name = "cppLabel";
            this.cppLabel.Size = new System.Drawing.Size(48, 16);
            this.cppLabel.TabIndex = 4;
            this.cppLabel.Text = "CPP：";
            // 
            // hourComb
            // 
            this.hourComb.FormattingEnabled = true;
            this.hourComb.Location = new System.Drawing.Point(75, 15);
            this.hourComb.Name = "hourComb";
            this.hourComb.Size = new System.Drawing.Size(46, 20);
            this.hourComb.TabIndex = 5;
            // 
            // dayComb
            // 
            this.dayComb.FormattingEnabled = true;
            this.dayComb.Location = new System.Drawing.Point(75, 48);
            this.dayComb.Name = "dayComb";
            this.dayComb.Size = new System.Drawing.Size(46, 20);
            this.dayComb.TabIndex = 6;
            
            // 
            // monthComb
            // 
            this.monthComb.FormattingEnabled = true;
            this.monthComb.Location = new System.Drawing.Point(75, 82);
            this.monthComb.Name = "monthComb";
            this.monthComb.Size = new System.Drawing.Size(46, 20);
            this.monthComb.TabIndex = 7;
            
            // 
            // yearComb
            // 
            this.yearComb.FormattingEnabled = true;
            this.yearComb.Location = new System.Drawing.Point(75, 118);
            this.yearComb.Name = "yearComb";
            this.yearComb.Size = new System.Drawing.Size(46, 20);
            this.yearComb.TabIndex = 8;
            
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(127, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(11, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "h";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(127, 51);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(23, 12);
            this.label6.TabIndex = 10;
            this.label6.Text = "day";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(127, 85);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 12);
            this.label7.TabIndex = 11;
            this.label7.Text = "months";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(127, 121);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(35, 12);
            this.label8.TabIndex = 12;
            this.label8.Text = "years";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(223, 110);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "确定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // CPPCalculate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(371, 191);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.yearComb);
            this.Controls.Add(this.monthComb);
            this.Controls.Add(this.dayComb);
            this.Controls.Add(this.hourComb);
            this.Controls.Add(this.cppLabel);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "CPPCalculate";
            this.Text = "CPPCalculate";
            this.Load += new System.EventHandler(this.CPPCalculate_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label cppLabel;
        private System.Windows.Forms.ComboBox hourComb;
        private System.Windows.Forms.ComboBox dayComb;
        private System.Windows.Forms.ComboBox monthComb;
        private System.Windows.Forms.ComboBox yearComb;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button1;
    }
}