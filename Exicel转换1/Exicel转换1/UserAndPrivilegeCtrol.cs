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
    public partial class UserAndPrivilegeCtrol : UserControl
    {
        UserAcount userAcount;
        public UserAndPrivilegeCtrol(UserAcount user)
        {
            InitializeComponent();
            //根据用户权限确认 并 显示

            ShowCtrolsBaseUserLevel(user.Level);
            userAcount = user;
        }



        //初始化显示
        private void ShowCtrolsBaseUserLevel(int userLevel)
        {
            if (userLevel!=0)
            {
                //禁用使用
                foreach (Control control in Controls)
                {
                    control.Enabled = false;
                }
            }
            
            //根据账户信息初始化
            switch (userLevel)
            {
                case 0://0, 最高权限
                    break;
                case 1://1, 权限管理用户
                    AllUserCombo.Enabled = true;
                    label5.Enabled = true;
                    label7.Enabled = true;
                    modefyPrivilegeCombo.Enabled = true;
                    ModefyBtn.Enabled = true;
                    break;
                case 2://2，专有管理员 可以编辑设定报价单的价格基准
                    break;
                case 3://3，管理员，可以查看编辑生成报价，可以查询sqlite中的log 价格的记录
                case 4://4, 普通用户（默认），进行评估报告的转换，不包含报价单CPP
                case 5://5, 查看已转换的评估报告信息  //除 2以上外均无权修改
                    break;
            }
        }

        

        private void UserAndPrivilegeCtrol_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> userNamePriviDict = userAcount.LoadAllUser();
            //LoadAllUser
            //初始化所有用户 ,及权限
            foreach (var item in userNamePriviDict)
            {
                AllUserCombo.Items.Add(item.Key);
                modefyPrivilegeCombo.Items.Add(item.Value);
                PrivilegeCombo.Items.Add(item.Value);
            }
            
            //禁止用户编辑 下拉列表
            
        }

        //修改用户
        private void ModefyBtn_Click(object sender, EventArgs e)
        {
            //判断是否修改了 再保存

            if (modefyUserNameTxt.Enabled)
            {
                if (!userAcount.ModefyUserName(AllUserCombo.Text, modefyUserNameTxt.Text))
                {
                    MessageBox.Show("用户名修改失败！");
                }
            }
            if (modefyPwdTxt.Enabled)
            {
                if (!userAcount.ModefyUserPwd(AllUserCombo.Text, modefyPwdTxt.Text))
                {
                    MessageBox.Show("密码修改失败！");
                }
            }
            if (modefyPrivilegeCombo.Enabled)
            {
                if (!userAcount.ModefyUserPriv(AllUserCombo.Text, Convert.ToInt32(modefyPrivilegeCombo.Text)))
                {
                    MessageBox.Show("用户权限修改失败！");
                }
            }
        }

        //用户名和对应权限的联动
        private void AllUserCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox allUserCombo = sender as ComboBox;
            if (allUserCombo == null)
            {
                return;
            }
            modefyPrivilegeCombo.SelectedItem = modefyPrivilegeCombo.Items[allUserCombo.SelectedIndex];
        }

        //添加用户
        private void AddUserBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(userNameTxt.Text) && !string.IsNullOrEmpty(pwdTxt.Text) &&
                        !string.IsNullOrEmpty(confirmPwdTxt.Text) && pwdTxt.Text.Trim() == confirmPwdTxt.Text.Trim() &&
                        !string.IsNullOrEmpty(PrivilegeCombo.Text))
                {
                    int createUserInt = UserAcount.CreatAcount(userNameTxt.Text.Trim(), pwdTxt.Text.Trim(), Convert.ToInt32(PrivilegeCombo.Text));
                    if (createUserInt==0)//0：失败，1：成功，2：用户名重复
                    {
                        MessageBox.Show("添加用户失败！");
                    }
                    if (createUserInt==1)
                    {
                        MessageBox.Show("添加用户成功！");
                    }
                    if (createUserInt==2)
                    {
                        MessageBox.Show("用户名重复，请修改后重试！");
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("添加用户失败！");
                return;
            }
        }
    }
}
