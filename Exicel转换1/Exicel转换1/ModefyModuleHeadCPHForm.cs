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
    public partial class ModefyModuleHeadCPHForm : Form
    {
        //全局唯一的单例
        public static ModefyModuleHeadCPHForm single_ModefyModuleHeadCPHForm = null;

        //发布事件，用于父窗体订阅 刷新数据
        public event EventHandler Update_module_Head_Cph_Base = null;

        //定一个标志位，是否修改了数据，修改则提示并进行保存
        static bool isModefiedData = false;

        //Module Head CPH 综合的module_head_cph_struct 结构
        //静态字段，用于父窗体获取 其值
        //list引用类型，直接保存修改数据即可
        private List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List;
        private Dictionary<string, int> base_StatisticsDict;
        //转换为单M的BaseCount
        private int baseCount_M_TotalM3;

        private ModefyModuleHeadCPHForm(List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List, 
            Dictionary<string, int> base_StatisticsDict)
        {
            InitializeComponent();
            this.module_Head_Cph_Structs_List = module_Head_Cph_Structs_List;
            InitModuleHeadCPHStructListToDataGridView();

            //base_StatisticsDict
            this.base_StatisticsDict = base_StatisticsDict;
            //base数量 *4 | *2
            this.baseCount_M_TotalM3 = base_StatisticsDict["4MBASE"] * 4 + base_StatisticsDict["2MBASE"]*2;
            //初始化Base的组合
            InitBaseCombination();
        }

        //单例模式的实现
        public static ModefyModuleHeadCPHForm Getsingle_ModefyModuleHeadCPHForm(List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List,
            Dictionary<string, int> base_StatisticsDict)
        {
            //每次打开获取显示ModefyModuleHeadCPHForm控件时，重置标识未
            isModefiedData = false;
            //已经生成，不在创建
            if (single_ModefyModuleHeadCPHForm != null)
            {
                return single_ModefyModuleHeadCPHForm;
            }
            return new ModefyModuleHeadCPHForm(module_Head_Cph_Structs_List, base_StatisticsDict);
        }

        //初始化更新Base的组合
        public void InitBaseCombination()
        {
            //
            combo_2M.Items.Clear();
            combo_4M.Items.Clear();

            combo_2M.Items.Add(base_StatisticsDict["2MBASE"].ToString());
            combo_4M.Items.Add(base_StatisticsDict["4MBASE"].ToString());
            //combo_2M.Items.

            //计算Base所有的组合            

            //最大2MBase个数
            int max2MBase = baseCount_M_TotalM3 / 2;
            //最小2MBase个数
            int min2MBase = baseCount_M_TotalM3 % 4;
            //最大4MBase个数
            int max4MBase = baseCount_M_TotalM3 / 4;
            //添加其他组合项目
            for (int i = min2MBase; i <= max2MBase; i++)
            {
                //排除现有项添加
                if (i== base_StatisticsDict["2MBASE"])
                {
                    continue;
                }
                //计算当前2MBase个数为i时，4M是否成立，不成立则更改
                if ((baseCount_M_TotalM3 - i * 2) % 4!=0)
                {
                    continue;
                }
                combo_2M.Items.Add(i);
                combo_4M.Items.Add((baseCount_M_TotalM3 - i * 2) / 4);
            }

            //combo_2M赋值
            combo_2M.Text = combo_2M.Items[0].ToString();
            combo_4M.Text = combo_4M.Items[0].ToString();

            //添加 2M base和4M Base的组合变化事件
            combo_2M.SelectedIndexChanged += Combo_2M_SelectedIndexChanged;
            combo_4M.SelectedIndexChanged += Combo_4M_SelectedIndexChanged;
        }

        //组合之间相互 indexchange改变的事件，不会相互之间执行形成死循环
        private void Combo_4M_SelectedIndexChanged(object sender, EventArgs e)
        {
            isModefiedData = true;
            ComboBox Combo_4M = sender as ComboBox;
            if (Combo_4M==null)
            {
                return;
            }
            //MessageBox.Show("我是4M改变了");
            combo_2M.SelectedItem = combo_2M.Items[Combo_4M.SelectedIndex];
        }

        private void Combo_2M_SelectedIndexChanged(object sender, EventArgs e)
        {
            //数据改动
            isModefiedData = true;
            ComboBox Combo_2M = sender as ComboBox;
            if (Combo_2M == null)
            {
                return;
            }
            //MessageBox.Show("我是2M改变了");
            combo_4M.SelectedItem = combo_4M.Items[Combo_2M.SelectedIndex];
        }
                
        //更新ComboBox组合的base数据到base字典
        private void UpadateComboBaseTobase_StatisticsDict()
        {
            base_StatisticsDict["4MBASE"] = Convert.ToInt32(combo_4M.Text);
            base_StatisticsDict["2MBASE"] = Convert.ToInt32(combo_2M.Text);
        }

        /// <summary>
        /// 将module_Head_Cph_Structs_List转换到ComboBoxColumn DataGridView
        /// </summary>
        /// <returns></returns>
        private void InitModuleHeadCPHStructListToDataGridView()
        {
            ModuleHeadCPHDGV.Columns.Add("Module", "Module");
            ModuleHeadCPHDGV.Columns.Add("Head", "Head");
            ModuleHeadCPHDGV.Columns.Add("TheoryCPH", "TheoryCPH");
            ModuleHeadCPHDGV.Columns.Add("placement", " ");

            //设置列宽 自动
            ModuleHeadCPHDGV.Columns["TheoryCPH"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //设置空余列的宽度
            ModuleHeadCPHDGV.Columns[3].Width = 150;
            //禁用空余列的编辑
            ModuleHeadCPHDGV.Columns[3].ReadOnly = true;

            //禁用列的自动排序
            for (int i = 0; i < ModuleHeadCPHDGV.Columns.Count; i++)
            {
                ModuleHeadCPHDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                ModuleHeadCPHDGV.Columns[i].DefaultCellStyle.Format = "c";
            }

            for (int i = 0; i < this.module_Head_Cph_Structs_List.Count; i++)
            {
                //添加新行
                ModuleHeadCPHDGV.Rows.Add();

                //添加module到comboBoxCell
                DataGridViewComboBoxCell moduledataGridViewComboBoxCell = new DataGridViewComboBoxCell();
                for (int j = 0; j < this.module_Head_Cph_Structs_List[i].moduleTrayTypestring.Length; j++)
                {
                    moduledataGridViewComboBoxCell.Items.Add(this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[j]);
                }

                //添加到单元格cell
                ModuleHeadCPHDGV.Rows[i].Cells["Module"] = moduledataGridViewComboBoxCell;

                //添加Head到comboBoxCell
                DataGridViewComboBoxCell headdataGridViewComboBoxCell = new DataGridViewComboBoxCell();
                //headdataGridViewComboBoxCell.Items.IndexOf;

                if (this.module_Head_Cph_Structs_List[i].headTypeString_List.Count != this.module_Head_Cph_Structs_List[i].CPH_strings_List.Count)
                {
                    MessageBox.Show("程序出现了严重问题，请不要使用本软件！");
                }

                for (int j = 0; j < this.module_Head_Cph_Structs_List[i].headTypeString_List.Count; j++)
                {
                    headdataGridViewComboBoxCell.Items.Add(this.module_Head_Cph_Structs_List[i].headTypeString_List[j]);

                    ////添加CPH到comboBoxCell,CPH的Cell必须和headtype一致，即head变化后CPH才会重新赋值
                    //DataGridViewComboBoxCell cphdataGridViewComboBoxCell = new DataGridViewComboBoxCell();
                    //for (int k = 0; k < this.module_Head_Cph_Structs_List[i].CPH_strings_List[j].Length; k++)
                    //{
                    //    cphdataGridViewComboBoxCell.Items.Add(this.module_Head_Cph_Structs_List[i].CPH_strings_List[j][k]);
                    //}

                }
                //添加当前i行 的Head列
                ModuleHeadCPHDGV.Rows[i].Cells["Head"] = headdataGridViewComboBoxCell;



            }

            //添加单元格内容修改时的事件，如果head列变化，cph列也变化
            ModuleHeadCPHDGV.CellValueChanged += ModuleHeadCPHDGV_CellValueChanged;

            //初始化 datagridview的cell的value
            //DataGridView返回的行数比实际行数 多1，最后一行无真实数据，转换报错（ModuleHeadCPHDGV.Count）
            for (int i = 0; i < ModuleHeadCPHDGV.Rows.Count-1; i++)
            {
                ModuleHeadCPHDGV.Rows[i].Cells["Module"].Value = (ModuleHeadCPHDGV.Rows[i].Cells["Module"] as DataGridViewComboBoxCell).Items[0];
                ModuleHeadCPHDGV.Rows[i].Cells["Head"].Value = (ModuleHeadCPHDGV.Rows[i].Cells["Head"] as DataGridViewComboBoxCell).Items[0];
            }
            //初始化过程中的数据改动 不应计算为修改了数据
            isModefiedData = false;
        }

        /// <summary>
        /// 保存DataGridView修改的数据到ModuleHeadCPHStructList
        /// </summary>
        private void DataGridViewToModuleHeadCPHStructList()
        {
            //DataGridView中实际数据为总Rows减1，最后一行无数据。ModuleHeadCPHDGV.Rows.Count-1
            for (int i = 0; i < ModuleHeadCPHDGV.Rows.Count-1; i++)
            {
                //ModuleHeadCPHDGV.Columns.Add("Module", "Module");
                //ModuleHeadCPHDGV.Columns.Add("Head", "Head");
                //ModuleHeadCPHDGV.Columns.Add("TheoryCPH", "TheoryCPH");
                //获取当前值
                string moduleValue = ModuleHeadCPHDGV.Rows[i].Cells["Module"].Value.ToString();
                string headValue = ModuleHeadCPHDGV.Rows[i].Cells["Head"].Value.ToString();
                string TheoryCPH = ModuleHeadCPHDGV.Rows[i].Cells["TheoryCPH"].Value.ToString();

                //修改module_Head_Cph_Structs_List[i]的数据
                for (int j = 0; j < this.module_Head_Cph_Structs_List[i].moduleTrayTypestring.Length; j++)
                {
                    //未改变对应的数据，即值仍未数组的第一个，则不进行下面的交换计算，省去开销
                    if (this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[0] == moduleValue)
                    {
                        break;
                    }
                    if (this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[j]== moduleValue)
                    {
                        //与数组的第一个交换位置
                        this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[j] = this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[0];
                        this.module_Head_Cph_Structs_List[i].moduleTrayTypestring[0] = moduleValue;
                    }
                }
                //headType、CPHStringsList
                for (int j = 0; j < this.module_Head_Cph_Structs_List[i].headTypeString_List.Count; j++)
                {
                    //未改变对应的数据，即值仍未数组的第一个，则不进行下面的交换计算，省去开销
                    if (this.module_Head_Cph_Structs_List[i].headTypeString_List[0] == headValue)
                    {
                        break;
                    }
                    if (this.module_Head_Cph_Structs_List[i].headTypeString_List[j] == headValue)
                    {
                        //改变数据后再进行交换，即值数值不是数组的第一个，才进行交换计算，省去开销
                        //if (j!=0)
                        //{
                            //与数组的第一个交换位置
                            this.module_Head_Cph_Structs_List[i].headTypeString_List[j] = this.module_Head_Cph_Structs_List[i].headTypeString_List[0];
                            this.module_Head_Cph_Structs_List[i].headTypeString_List[0] = headValue;

                            //交换CPH的List
                            string[] CPH_strings_List = this.module_Head_Cph_Structs_List[i].CPH_strings_List[j];
                            this.module_Head_Cph_Structs_List[i].CPH_strings_List[j] = this.module_Head_Cph_Structs_List[i].CPH_strings_List[0];
                            this.module_Head_Cph_Structs_List[i].CPH_strings_List[0] = CPH_strings_List;
                        //}               
                    }
                }


                //交换CPHStringsList内的strings(即 TheoryCPH的改动)，这个改动是CPHStringsList内的string[]数组的交换，
                //独立于//headType、CPHStringsList
                for (int k = 0; k < this.module_Head_Cph_Structs_List[i].CPH_strings_List[0].Length; k++)
                {
                    if (this.module_Head_Cph_Structs_List[i].CPH_strings_List[0][k] == TheoryCPH)
                    {
                        this.module_Head_Cph_Structs_List[i].CPH_strings_List[0][k] = this.module_Head_Cph_Structs_List[i].CPH_strings_List[0][0];
                        this.module_Head_Cph_Structs_List[i].CPH_strings_List[0][0] = TheoryCPH;
                    }
                }
            }
        }

        private void ModuleHeadCPHDGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == null)
            {
                return;
            }
            //修改标志位，数据改动
            isModefiedData = true;
            //Head与对应的Theory CPH的联动实现。
            if (e.ColumnIndex == 1)//Head列时处理
            {
                //head选择的第k个
                int k = 0;

                DataGridViewComboBoxCell gridViewComboBoxCell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewComboBoxCell;
                for (int i = 0; i < gridViewComboBoxCell.Items.Count; i++)
                {
                    if (gridViewComboBoxCell.Items[i].ToString() == gridViewComboBoxCell.Value.ToString())
                    {
                        k = i;
                    }
                }

                //CPH对应也是第k个
                DataGridViewComboBoxCell cphgridViewComboBoxCell = new DataGridViewComboBoxCell();
                for (int i = 0; i < this.module_Head_Cph_Structs_List[e.RowIndex].CPH_strings_List[k].Length; i++)
                {
                    cphgridViewComboBoxCell.Items.Add(this.module_Head_Cph_Structs_List[e.RowIndex].CPH_strings_List[k][i]);
                }

                dataGridView.Rows[e.RowIndex].Cells["TheoryCPH"] = cphgridViewComboBoxCell;
                dataGridView.Rows[e.RowIndex].Cells["TheoryCPH"].Value = cphgridViewComboBoxCell.Items[0];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (isModefiedData)
            {
                if (MessageBox.Show("修改了数值，请确认是否进行保存！", "确认!", MessageBoxButtons.YesNo)!=DialogResult.Yes)
                {//不确认更改，取消后面的保存
                    return;
                }
                

            }

            //保存修改的DataDridView数据
            DataGridViewToModuleHeadCPHStructList();

            //保存2M、4M Base组合到字典
            UpadateComboBaseTobase_StatisticsDict();

            //重新初始化Base的combBox控件内容
            //实际使用中发现，选择更改后，combBox在再一次打开后没有多余选项，或者选项重复，索性一选择重新初始化
            //并且重写InitBaseCombination，先清空在添加，然后在添加设置
            InitBaseCombination();

            //可简化委托调用Update_module_Head_Cph_Base?.Invoke(this, null);
            if (Update_module_Head_Cph_Base!=null)//触发执行事件
            {
                Update_module_Head_Cph_Base(this, null);
            }

            
            //保存结束重置标志位
            isModefiedData = false;
            //关闭窗口
            this.Close();
        }
    }
}
