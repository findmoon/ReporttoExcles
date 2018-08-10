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
    public partial class LoginFrom : Form
    {
        LoginCtrol loginCtrol;
        RegisterCtrol registerCtrol;
        public LoginFrom()
        {
            InitializeComponent();
            //初始化登录和注册用户控件
            loginCtrol = new LoginCtrol();
            loginCtrol.LoginSuccess += LoginCtrol_LoginSuccess;//注册登录成功的事件
            registerCtrol = new RegisterCtrol();
            registerCtrol.RegisterSuccess += RegisterCtrol_RegisterSuccess;
        }

        //注册成功，显示登录界面
        private void RegisterCtrol_RegisterSuccess(object sender, EventArgs e)
        {
            ShowLogin();
        }

        //登录成功事件的处理方法
        private void LoginCtrol_LoginSuccess(object sender, EventArgs e)
        {
            Hide();
            //将登录的账户传递给主窗体
            XmlExcelRtfConvertMain_Form mainForm = new XmlExcelRtfConvertMain_Form(loginCtrol.userAcount);            
            mainForm.Show();
            //主窗体关闭事件，主窗体关闭仍在运行，登陆窗体未关闭
            mainForm.FormClosed += MainForm_FormClosed;
            //关闭当前窗体
            //Close();//关闭登陆窗体，并且释放窗体内的资源。实际关闭了所有的整个窗体，原因？
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkLabel1.Text.Trim() == "注册")
            {
                ShowRegister();
            }
            else
            {
                ShowLogin();
            }
        }

        private void LoginFrom_Load(object sender, EventArgs e)
        {
            //初始化时显示登录
            ShowLogin();
        }

        private void ShowLogin()
        {
            //显示登录
            linkLabel1.Text = "注册";
            loginRegisterGroB.Controls.Clear();
            loginRegisterGroB.Controls.Add(loginCtrol);
            registerCtrol.Enabled = false;
            loginCtrol.Enabled = true;
            loginCtrol.Show();
        }

        private void ShowRegister()
        {
            linkLabel1.Text = "返回登录";
            loginRegisterGroB.Controls.Clear();
            loginRegisterGroB.Controls.Add(registerCtrol);
            loginCtrol.Enabled = false;
            registerCtrol.Enabled = true;
            registerCtrol.Show();
        }

        private void LoginFrom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                e.Handled = true;   //将Handled设置为true，指示已经处理过KeyPress事件
                if (loginCtrol.Enabled)
                {
                    loginCtrol.LoginBtn.PerformClick();
                }
                if (registerCtrol.Enabled)
                {
                    registerCtrol.RegisterBtn.PerformClick();
                }                
            }
        }
    }
}
