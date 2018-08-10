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
    public partial class CurrenUerCtrol : UserControl
    {
        //标识位，用户名是否改变
        bool userNameHasChanged = false;
        UserAcount userAcount;
        public CurrenUerCtrol(UserAcount userAcount)
        {
            InitializeComponent();
            oldpwdTxt.PasswordChar = '*';
            newpwdTxt.PasswordChar = '*';
            confirmNewpwdTxt.PasswordChar = '*';

            userNameTxt.Text = userAcount.UserName;
            OpenEditFunc(false);
        }

        //启用和打开编辑功能
        public void OpenEditFunc(bool isopen)
        {            
                oldpwdTxt.Enabled= isopen;
                newpwdTxt.Enabled = isopen;
                confirmNewpwdTxt.Enabled = isopen;
                userNameTxt.Enabled = isopen;
            SaveBtn.Enabled = isopen;
        }

        private void CurrenUerCtrol_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkLabel1.Text=="编辑")
            {
                linkLabel1.Text = "取消编辑";
                OpenEditFunc(true);
            }
            else
            {
                linkLabel1.Text = "编辑";
                OpenEditFunc(false);
            }
        }

        ////检查是否可以修改密码，即密码正确，且新密码确认一致
        //private bool CanModifyPwd()
        //{
        //    oldpwdTxt.PasswordChar = '*';
        //    .PasswordChar = '*';
        //    .PasswordChar = '*';
        //}

        //保存功能
        private void SaveBtn_Click(object sender, EventArgs e)
        {
            //确认是否执行密码修改
            if (!string.IsNullOrEmpty(oldpwdTxt.Text) && !string.IsNullOrEmpty(newpwdTxt.Text) &&
                !string.IsNullOrEmpty(confirmNewpwdTxt.Text) && newpwdTxt.Text.Trim()== confirmNewpwdTxt.Text.Trim())
            {
                userAcount.ModefyUserPwd(userAcount.UserName, oldpwdTxt.Text.Trim(), newpwdTxt.Text.Trim());
            }
            else
            {
                if (userNameHasChanged)
                {
                    userAcount.ModefyUserName(userAcount.UserName, userNameTxt.Text.Trim());
                }
                else
                {
                    MessageBox.Show("没有要保存的信息！");
                }
            }
        }

        //用户名变化
        private void userNameTxt_TextChanged(object sender, EventArgs e)
        {
            userNameHasChanged = true;
        }
    }
}
