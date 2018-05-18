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
    //声明委托
    public delegate void sendValue_del(string productionModle);
    public partial class productionModleFrom : Form
    {
        //传值的变量
        public string production;

        //声明或创建委托变量
        public sendValue_del _sendValue;
        //public productionModleFrom(sendValue_del sendValue)
        //{
        //    InitializeComponent();
        //    //委托变量赋值
        //    this._sendValue = sendValue;
        //}

        public productionModleFrom()
        {
            InitializeComponent();
            this.radioButton1.Checked=true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ////执行委托

            //foreach (Control c in groupBox1.Controls)
            //{
            //    if ((c is RadioButton) && (c as RadioButton).Checked)
            //    {
            //        //执行委托
            //        this._sendValue(c.Text);
            //    }
            //}
            //this.Hide();

            ////将当前窗体拥有者强制转换为XmlReport类的实例
            //XmlReport xmlReport = (XmlReport)this.Owner;
            ////窗体的显示类型转换
            //foreach (Control c in groupBox1.Controls)
            //{
            //    if ((c is RadioButton) && (c as RadioButton).Checked)
            //    {
            //        xmlReport.ProductionMode= c.Text;
            //    }
            //}
            foreach (Control c in groupBox1.Controls)
            {
                if ((c is RadioButton) && (c as RadioButton).Checked)
                {
                    this.production = c.Text;
                }
            }
            this.DialogResult = DialogResult.OK;//将productionModleFrom属性DialogResult设置为ok
            this.Hide();
        }
    }
}
