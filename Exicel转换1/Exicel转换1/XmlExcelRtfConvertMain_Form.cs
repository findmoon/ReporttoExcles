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
        public UserManageCtrol userManage = null;
        public UserAcount userAcount;

        public XmlExcelRtfConvertMain_Form(UserAcount user)
        {
            InitializeComponent();
            //根据userAcount初始化Main窗体
            userAcount = user;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            ////根据userAcount初始化 生成评估报告用户控件【显示和限制不同功能】及用户管理模块
            userManage = new UserManageCtrol(userAcount);

            //控制log按钮的显示
            if (userAcount.Level > 3)
            {
                DetialLogBtn.Enabled = false;
            }
            else
            {
                DetialLogBtn.Enabled = false;
            }

            lineCfg = new LineCfg();//给窗体变量初始化
            flexCVCReport = new FlexCVReport();//窗体变量初始化
            xmlReport = new XmlReport(userAcount);//窗体变量初始化

            ////设置控件依靠父控件
            //xmlReport.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom);
            //xmlReport.Dock = DockStyle.Fill;

            xmlReport.Show();//显示xmlreport窗体控件
            XmlWindowToolBtn.BackColor = SystemColors.GradientActiveCaption;
            GpbWindow.Controls.Add(xmlReport);//加载xmlreport窗体控件

            //禁用和隐藏两个按钮
            FlexCVWindowToolBtn.Enabled = false;
            FlexCVWindowToolBtn.Visible = false;
            LinCfgWindowToolBtn.Enabled = false;
            LinCfgWindowToolBtn.Visible = false;


        }
        private void GpbWindow_Enter(object sender, EventArgs e)
        {
            //Dock 获取或设置哪些控件边框停靠到其父控件并确定控件如何随其父级一起调整大小。
            this.GpbWindow.AutoSizeMode = AutoSizeMode.GrowAndShrink;

        }

        private void XmlWindowToolBtn_Click(object sender, EventArgs e)
        {
            xmlReport.Show();//显示xmlreport窗体控件
            //XmlWindowToolBtn 颜色当前选中
            //XmlWindowToolBtn.BackColor = SystemColors.GradientInactiveCaption;
            XmlWindowToolBtn.BackColor = SystemColors.GradientActiveCaption;
            flexCVCReport.BackColor = SystemColors.ControlLight;
            lineCfg.BackColor = SystemColors.ControlLight;
            //临时隐藏
            flexCVCReport.Hide();
            lineCfg.Hide();
            GpbWindow.Controls.Clear();//清空之前加载的窗体控件
            GpbWindow.Controls.Add(xmlReport);//加载xmlreport窗体控件
        }

        private void FlexCVWindowToolBtn_Click(object sender, EventArgs e)
        {
            flexCVCReport.Show();

            flexCVCReport.BackColor = SystemColors.GradientActiveCaption;
            XmlWindowToolBtn.BackColor = SystemColors.ControlLight;
            lineCfg.BackColor = SystemColors.ControlLight;
            //临时
            xmlReport.Hide();
            lineCfg.Hide();
            GpbWindow.Controls.Clear();
            GpbWindow.Controls.Add(flexCVCReport);
            //输出日志
        }

        private void LinCfgWindowToolBtn_Click(object sender, EventArgs e)
        {
            lineCfg.Show();
            lineCfg.BackColor = SystemColors.GradientActiveCaption;
            XmlWindowToolBtn.BackColor = SystemColors.ControlLight;
            flexCVCReport.BackColor = SystemColors.ControlLight;
            //临时
            xmlReport.Hide();
            flexCVCReport.Hide();
            GpbWindow.Controls.Clear();
            GpbWindow.Controls.Add(lineCfg);
        }

        private void DetialLogBtn_Click(object sender, EventArgs e)
        {

        }

        private void UserAdmin_Click(object sender, EventArgs e)
        {
            GpbWindow.Controls.Clear();
            GpbWindow.Controls.Add(userManage);
            userManage.Show();
        }
    }
}
