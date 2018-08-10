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
    public partial class LoginCtrol : UserControl
    {
        //登录成功事件
        public event EventHandler LoginSuccess=null;

        //能够实现，在登录成功事件中传递 用户账户对象
        public UserAcount userAcount=null;

        public LoginCtrol()
        {
            InitializeComponent();
            pwdTxt.PasswordChar = '*';
        }

        private void LoginBtn_Click(object sender, EventArgs e)
        {
            //UserAcount的登录方法 实现登陆
            userAcount = UserAcount.LoginAcount(userTxt.Text.Trim(), pwdTxt.Text.Trim());
            if (userAcount!=null)//登录成功
            {
                if (userAcount.accountState==0)
                {
                    MessageBox.Show("密码过期！");
                    return;
                }
                if (userAcount.accountState==2)
                {
                    MessageBox.Show("登录失败，用户名或密码错误！");
                    return;
                }
                if (userAcount.accountState==1)
                {
                    //登录成功后，用户控件的父级窗体要关闭，显示出主窗体
                    //使用事件实现，触发事件
                    LoginSuccess?.Invoke(this, null);
                }
                                
            }
            else
            {
                MessageBox.Show("无法登陆，请重试！");
                return;
            }
        }
    }
}
