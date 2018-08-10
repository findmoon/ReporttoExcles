namespace Exicel转换1
{
    partial class QuotationForm
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
            this.QuotationDataGV = new System.Windows.Forms.DataGridView();
            this.discountComboB = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.discountCheckB = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.feederNozzlegroupB = new System.Windows.Forms.GroupBox();
            this.NozzleQtyComboB = new System.Windows.Forms.ComboBox();
            this.FeederQtyComboB = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.feederNozzleQtyCheckB = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.GenQuotationExcelBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.QuotationDataGV)).BeginInit();
            this.feederNozzlegroupB.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // QuotationDataGV
            // 
            this.QuotationDataGV.AllowUserToAddRows = false;
            this.QuotationDataGV.AllowUserToDeleteRows = false;
            this.QuotationDataGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.QuotationDataGV.Location = new System.Drawing.Point(6, 80);
            this.QuotationDataGV.Name = "QuotationDataGV";
            this.QuotationDataGV.ReadOnly = true;
            this.QuotationDataGV.RowTemplate.Height = 23;
            this.QuotationDataGV.Size = new System.Drawing.Size(828, 454);
            this.QuotationDataGV.TabIndex = 0;
            // 
            // discountComboB
            // 
            this.discountComboB.FormattingEnabled = true;
            this.discountComboB.Items.AddRange(new object[] {
            "10",
            "11",
            "12",
            "13",
            "14",
            "15"});
            this.discountComboB.Location = new System.Drawing.Point(100, 20);
            this.discountComboB.Name = "discountComboB";
            this.discountComboB.Size = new System.Drawing.Size(62, 20);
            this.discountComboB.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(168, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "折";
            // 
            // discountCheckB
            // 
            this.discountCheckB.AutoSize = true;
            this.discountCheckB.Location = new System.Drawing.Point(46, 22);
            this.discountCheckB.Name = "discountCheckB";
            this.discountCheckB.Size = new System.Drawing.Size(48, 16);
            this.discountCheckB.TabIndex = 4;
            this.discountCheckB.Text = "折扣";
            this.discountCheckB.UseVisualStyleBackColor = true;
            this.discountCheckB.CheckedChanged += new System.EventHandler(this.discountCheckB_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(3, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(210, 14);
            this.label2.TabIndex = 5;
            this.label2.Text = "注意：不包含FLEXA等软体价格！";
            // 
            // feederNozzlegroupB
            // 
            this.feederNozzlegroupB.Controls.Add(this.NozzleQtyComboB);
            this.feederNozzlegroupB.Controls.Add(this.FeederQtyComboB);
            this.feederNozzlegroupB.Controls.Add(this.label4);
            this.feederNozzlegroupB.Controls.Add(this.label3);
            this.feederNozzlegroupB.Controls.Add(this.feederNozzleQtyCheckB);
            this.feederNozzlegroupB.Location = new System.Drawing.Point(234, 6);
            this.feederNozzlegroupB.Name = "feederNozzlegroupB";
            this.feederNozzlegroupB.Size = new System.Drawing.Size(275, 57);
            this.feederNozzlegroupB.TabIndex = 6;
            this.feederNozzlegroupB.TabStop = false;
            // 
            // NozzleQtyComboB
            // 
            this.NozzleQtyComboB.FormattingEnabled = true;
            this.NozzleQtyComboB.Items.AddRange(new object[] {
            "1",
            "1.1",
            "1.2",
            "1.3",
            "1.4",
            "1.5",
            "1.6",
            "1.7",
            "1.8",
            "1.9",
            "2.0"});
            this.NozzleQtyComboB.Location = new System.Drawing.Point(180, 37);
            this.NozzleQtyComboB.Name = "NozzleQtyComboB";
            this.NozzleQtyComboB.Size = new System.Drawing.Size(55, 20);
            this.NozzleQtyComboB.TabIndex = 4;
            // 
            // FeederQtyComboB
            // 
            this.FeederQtyComboB.FormattingEnabled = true;
            this.FeederQtyComboB.Items.AddRange(new object[] {
            "1",
            "1.1",
            "1.2",
            "1.3",
            "1.4",
            "1.5",
            "1.6",
            "1.7",
            "1.8",
            "1.9",
            "2.0"});
            this.FeederQtyComboB.Location = new System.Drawing.Point(180, 7);
            this.FeederQtyComboB.Name = "FeederQtyComboB";
            this.FeederQtyComboB.Size = new System.Drawing.Size(55, 20);
            this.FeederQtyComboB.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(108, 42);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "Nozzle Qty：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(108, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "Feeder Qty：";
            // 
            // feederNozzleQtyCheckB
            // 
            this.feederNozzleQtyCheckB.AutoSize = true;
            this.feederNozzleQtyCheckB.Location = new System.Drawing.Point(0, 13);
            this.feederNozzleQtyCheckB.Name = "feederNozzleQtyCheckB";
            this.feederNozzleQtyCheckB.Size = new System.Drawing.Size(102, 16);
            this.feederNozzleQtyCheckB.TabIndex = 0;
            this.feederNozzleQtyCheckB.Text = "Feeder/Nozzle";
            this.feederNozzleQtyCheckB.UseVisualStyleBackColor = true;
            this.feederNozzleQtyCheckB.CheckedChanged += new System.EventHandler(this.feederNozzleQtyCheckB_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.discountComboB);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.discountCheckB);
            this.groupBox2.Location = new System.Drawing.Point(555, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(205, 57);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(473, 12);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(17, 12);
            this.label5.TabIndex = 5;
            this.label5.Text = "倍";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(473, 41);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 5;
            this.label6.Text = "倍";
            // 
            // GenQuotationExcelBtn
            // 
            this.GenQuotationExcelBtn.Location = new System.Drawing.Point(79, 32);
            this.GenQuotationExcelBtn.Name = "GenQuotationExcelBtn";
            this.GenQuotationExcelBtn.Size = new System.Drawing.Size(82, 31);
            this.GenQuotationExcelBtn.TabIndex = 5;
            this.GenQuotationExcelBtn.Text = "生成报价单";
            this.GenQuotationExcelBtn.UseVisualStyleBackColor = true;
            this.GenQuotationExcelBtn.Click += new System.EventHandler(this.GenQuotationExcelBtn_Click);
            // 
            // QuotationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(846, 543);
            this.Controls.Add(this.GenQuotationExcelBtn);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.feederNozzlegroupB);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.QuotationDataGV);
            this.Name = "QuotationForm";
            this.Text = "QuotationForm";
            this.Load += new System.EventHandler(this.QotationForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.QuotationDataGV)).EndInit();
            this.feederNozzlegroupB.ResumeLayout(false);
            this.feederNozzlegroupB.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.DataGridView QuotationDataGV;
        private System.Windows.Forms.ComboBox discountComboB;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox discountCheckB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox feederNozzlegroupB;
        private System.Windows.Forms.ComboBox NozzleQtyComboB;
        private System.Windows.Forms.ComboBox FeederQtyComboB;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox feederNozzleQtyCheckB;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button GenQuotationExcelBtn;
    }
}