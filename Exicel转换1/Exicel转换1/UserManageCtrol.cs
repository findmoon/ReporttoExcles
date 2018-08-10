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
    public partial class UserManageCtrol : UserControl
    {
        public UserManageCtrol(UserAcount userAcount)
        {
            InitializeComponent();

            UserAdminTab.TabPages.Add("CurrentUser", "当前用户");
            UserAdminTab.TabPages.Add("UserPriv", "用户管理");
            //当前用户page
            UserAdminTab.TabPages[0].Controls.Add(new CurrenUerCtrol(userAcount));
            //管理 page
            UserAdminTab.TabPages[1].Controls.Add(new UserAndPrivilegeCtrol(userAcount));
            ////根据账户信息初始化tab
            //switch (userAcount.Level)
            //{
            //    case 0://最高权限
            //        break;
            //    case 1://1, 权限管理用户
            //        break;
            //    case 2://2，专有管理员 可以编辑设定报价单的价格基准
            //        break;
            //    case 3://3，管理员，可以查看编辑生成报价，可以查询sqlite中的log 价格的记录
            //        break;
            //    case 4://4, 普通用户（默认），进行评估报告的转换，不包含报价单CPP

            //        break;
            //    case 5://5, 查看已转换的评估报告信息
            //    default:
            //        break;
            //}
            //UserAdminTab.
        }

        private void UserManageCtrol_Load(object sender, EventArgs e)
        {
            #region 标签左侧显示
            //重绘tabctrol，
            UserAdminTab.DrawMode = TabDrawMode.OwnerDrawFixed;//设置为用户绘制模式
            UserAdminTab.Alignment = TabAlignment.Left;//选项卡放到左边
            UserAdminTab.SizeMode = TabSizeMode.Fixed;//设定所有选显卡具有相同大小
            UserAdminTab.ItemSize = new Size(50, 120);//设置选项卡大小 
            #endregion
        }

        private void UserAdminTab_DrawItem(object sender, DrawItemEventArgs e)
        {
            Font fntTab;
            Brush brushBack, brushFore;
            if (e.Index == ((TabControl)(sender)).SelectedIndex)//当前Tab页面的样式
            {
                fntTab = new Font(e.Font, FontStyle.Bold);
                //brushBack = new System.Drawing.Drawing2D.LinearGradientBrush(e.Bounds, SystemColors.Control, SystemColors.ControlLightLight, System.Drawing.Drawing2D.LinearGradientMode.Horizontal);
                brushBack = new SolidBrush(Color.LightSlateGray);
                brushFore = Brushes.Black;
            }
            else//其他tab样式
            {
                fntTab = e.Font;
                brushBack = new SolidBrush(SystemColors.Control);
                brushFore = new SolidBrush(Color.Black);
            }
            //画样式
            e.Graphics.FillRectangle(brushBack, e.Bounds);
            string tabName = ((TabControl)(sender)).TabPages[e.Index].Text;
            StringFormat sfTab = new StringFormat();
            sfTab.LineAlignment = StringAlignment.Center;//垂直对齐，指定文本在布局矩形的中心对齐
            sfTab.Alignment = StringAlignment.Center;//水平对齐
            Rectangle recTab = e.Bounds;
            recTab = new Rectangle(recTab.X, recTab.Y + 4, recTab.Width, recTab.Height - 4);
            e.Graphics.DrawString(tabName, fntTab, brushFore, recTab, sfTab);
        }
    }
}
