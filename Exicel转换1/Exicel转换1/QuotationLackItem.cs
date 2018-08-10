using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    public partial class QuotationLackItem : Form
    {
        public QuotationLackItem(string lackString)
        {
            InitializeComponent();
            lackItemTxt.Text = lackString;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }
    }
}
