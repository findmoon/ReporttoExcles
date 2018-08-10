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
    public partial class CPPCalculate : Form
    {
        //使用自定义类的泛型委托
        public event EventHandler<CppEventArgs> GetCPPEvent=null;        
        double cpp;
        //public double Cpp { get { return cpp; } }

        double minPrice;
        int cph;
        public CPPCalculate(double minPrice,int cph)
        {
            InitializeComponent();
            this.minPrice = minPrice;
            this.cph = cph;
        }          

        private void CPPCalculate_Load(object sender, EventArgs e)
        {
            //初始化combBox
            object[] hours = new object[5]{ 22, 23, 24, 21, 20 };      
            hourComb.Items.AddRange(hours);
            hourComb.Text = hourComb.Items[0].ToString();

            object[] days = new object[5] { 26, 27, 28, 29, 30 };
            dayComb.Items.AddRange(days);
            dayComb.Text = dayComb.Items[0].ToString();

            object[] mouths = new object[2] { 12, 11 };
            monthComb.Items.AddRange(mouths);
            monthComb.Text = monthComb.Items[0].ToString();

            object[] years = new object[5] { 5, 4, 3, 2, 1 };
            yearComb.Items.AddRange(years);
            yearComb.Text = yearComb.Items[0].ToString();

            //初始化之后 订阅事件
            yearComb.TextChanged += new EventHandler(yearComb_TextChanged);
            monthComb.TextChanged += new EventHandler(monthComb_TextChanged);
            dayComb.TextChanged += new EventHandler(dayComb_TextChanged);
            hourComb.TextChanged += HourComb_TextChanged;
            CalculateCPP();
        }

        void CalculateCPP()
        {
            try
            {
                cpp = minPrice / Convert.ToDouble(hourComb.Text) / Convert.ToDouble(dayComb.Text) /
                Convert.ToDouble(monthComb.Text) / Convert.ToDouble(yearComb.Text)/ Convert.ToDouble(cph);
                cppLabel.Text = "CPP：" + cpp;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void HourComb_TextChanged(object sender, EventArgs e)
        {
            CalculateCPP();            
        }
        private void dayComb_TextChanged(object sender, EventArgs e)
        {
            CalculateCPP();
        }

        private void monthComb_TextChanged(object sender, EventArgs e)
        {
            CalculateCPP();
        }

        private void yearComb_TextChanged(object sender, EventArgs e)
        {
            CalculateCPP();
        }

        //CPP确定事件，触发传送CPP的值
        private void button1_Click(object sender, EventArgs e)
        {
            CppEventArgs args = new CppEventArgs() { CPP = cpp };
            GetCPPEvent?.Invoke(this, args);
            Hide();
        }
    }

    //派生EventArgs，存储床底的cpp数据
    public class CppEventArgs : EventArgs
    {
        public double CPP { get; set; }
    }
}
