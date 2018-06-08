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
    public partial class XmlExcelRtfConvertMain_Form : Form
    {
        public LineCfg lineCfg;
        public FlexCVReport flexCVCReport;
        public XmlReport xmlReport;

        public XmlExcelRtfConvertMain_Form()
        {
            InitializeComponent();
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            lineCfg = new LineCfg();//给窗体变量初始化
            flexCVCReport = new FlexCVReport();//窗体变量初始化
            xmlReport = new XmlReport();//窗体变量初始化

            xmlReport.Show();//显示xmlreport窗体控件
            GpbWindow.Controls.Add(xmlReport);//加载xmlreport窗体控件

            //禁用和隐藏两个按钮
            FlexCVWindow.Enabled = false;
            FlexCVWindow.Visible = false;
            LinCfgWindow.Enabled = false;
            LinCfgWindow.Visible = false;
        }

        private void XmlWindow_Click(object sender, EventArgs e)
        {
            xmlReport.Show();//显示xmlreport窗体控件
            flexCVCReport.Hide();
            lineCfg.Hide();
            GpbWindow.Controls.Clear();//清空之前加载的窗体控件
            GpbWindow.Controls.Add(xmlReport);//加载xmlreport窗体控件
            //输出日志
        }

        private void FlexCVWindow_Click(object sender, EventArgs e)
        {
            flexCVCReport.Show();
            xmlReport.Hide();
            lineCfg.Hide();
            GpbWindow.Controls.Clear();
            GpbWindow.Controls.Add(flexCVCReport);
            //输出日志
        }

        private void LinCfgWindow_Click(object sender, EventArgs e)
        {
            lineCfg.Show();
            xmlReport.Hide();
            flexCVCReport.Hide();
            GpbWindow.Controls.Clear();
            GpbWindow.Controls.Add(lineCfg);
        }       
    }
}
