using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    public partial class RegisterCtrol : UserControl
    {
        //登录成功事件
        public event EventHandler RegisterSuccess = null;
        public RegisterCtrol()
        {
            InitializeComponent();
            pwdTxt.PasswordChar = '*';
            confirmPwdtxt.PasswordChar = '*';
        }

        private void ClearInfo()
        {
            //清空用户名等
            userTxt.Text = string.Empty;
            pwdTxt.Text = string.Empty;
            confirmPwdtxt.Text = string.Empty;
        }

        private void RegisterBtn_Click(object sender, EventArgs e)
        {
            int creatUserInt = UserAcount.CreatAcount(userTxt.Text.Trim(), pwdTxt.Text.Trim());
            //注册是否成功的判断。0：失败，1：成功，2：用户名重复
            switch (creatUserInt)
            {
                case 0:
                    MessageBox.Show("注册失败！请重试！");
                    break;
                case 1:
                    //触发注册成功事件
                    MessageBox.Show("注册成功！");
                    ClearInfo();
                    RegisterSuccess?.Invoke(this, null);
                    break;
                case 2:
                    MessageBox.Show("用户名已存在！");
                    ClearInfo();
                    break;
                default:
                    break;
            }
        }
        //验证是否相等
        private void confirmPwdtxt_TextChanged(object sender, EventArgs e)
        {
            if (confirmPwdtxt.Text.Trim()!=pwdTxt.Text.Trim())
            {
                errorProvider1.SetError(confirmPwdtxt, "密码不一致!");
            }
        }

        private void pwdTxt_TextChanged(object sender, EventArgs e)
        {
            //验证密码和确认密码的正则
            //string patternStr = "(a-z|0-9)|(A-Z|0-9)|";
            if (true)
            {

            }
        }
    }
}
