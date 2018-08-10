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
    public partial class Single_EvaluationPanelControl : UserControl
    {
        //保存数据的方式：
        //1、通过传递过来的baseComprehensive对象，直接执行其更新方法，保存数据
        //2、通过发布事件，触发执行事件。XmlReport对象中订阅事件，事件方法中执行当前baseComprehensive对象的更新方法
        //第二种，当有多个baseComprehensive对象时，事件方法可通过当前TabPage索引或对象TimeStamp字段确认更新哪个
        private Int64 timeStamp;
        private string jobName;
        private List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List;
        private Dictionary<string, int> base_StatisticsDict;
        ModefyModuleHeadCPHForm modefyModuleHeadCPHForm = null;
        //执行更新所需的BaseComprehensive baseComprehensive
        EvaluationReportClass evaluationReportClass = null;

        //当前的报价单对象
        QuotationForm qotationForm = null;

        //public double CPP { get; set; }

        public Int64 TimeStamp
        {
            get
            {
                return this.timeStamp;
            }
        }
        public string JobNmae
        {
            get
            {
                return this.jobName;
            }
        }
        /// <summary>
        /// 初始化 Single_EvaluationPanel
        /// </summary>
        /// <param name="baseComprehensive">BaseComprehensive对象 作为参数传递</param>
        /// <param name="userLevel">用户权限等级 作为参数传递</param>
        ///public Single_EvaluationPanelControl(Int64 timeStamp,string jobName, DataTable summaryDT,DataTable layoutDT,
        ///    DataTable feederNozzleDT, Dictionary<string, int> base_StatisticsDict, ref List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List)
        public Single_EvaluationPanelControl(EvaluationReportClass baseComprehensive, int userLevel)
        {
            InitializeComponent();

            //是否启用编辑报价单
            if (userLevel == 3 || userLevel == 0)
            {
                ViewEditQotation.Enabled = true;
            }
            else
            {
                ViewEditQotation.Enabled = false;
            }

            //baseComprehensive.TimeStamp, 
            //baseComprehensive.Jobname, baseComprehensive.SummaryInfoDT,baseComprehensive.layoutDT,
            //baseComprehensive.feederNozzleDT, baseComprehensive.base_StatisticsDict,
            //ref baseComprehensive.module_Head_Cph_Structs_List
            this.evaluationReportClass = baseComprehensive;
            timeStamp = baseComprehensive.TimeStamp;
            jobName = baseComprehensive.Jobname;
            summaryDGV.DataSource = baseComprehensive.SummaryDT;
            layoutDGV.DataSource = baseComprehensive.layoutDT;
            feederNozzleDGV.DataSource = baseComprehensive.feederNozzleDT;

            #region DataGridView的一些设置
            //数字的显示格式是￥12
            summaryDGV.DefaultCellStyle.Format = "G";//均正常显示
            summaryDGV.Columns[11].DefaultCellStyle.Format = "P";//cphrate 百分比
            layoutDGV.DefaultCellStyle.Format = "G";//均正常显示
            feederNozzleDGV.DefaultCellStyle.Format = "G";//均正常显示

            module_Head_Cph_Structs_List = baseComprehensive.module_Head_Cph_Structs_List;
            base_StatisticsDict = baseComprehensive.base_StatisticsDict;

            //不显示行标题
            this.summaryDGV.RowHeadersVisible = false;
            //this.summaryDGV.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.DisableResizing);
            this.layoutDGV.RowHeadersVisible = false;
            this.feederNozzleDGV.RowHeadersVisible = false;

            //禁用列的自动排序
            for (int i = 0; i < summaryDGV.Columns.Count; i++)
            {
                this.summaryDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.summaryDGV.Columns[i].DefaultCellStyle.Format = "c";
            }
            for (int i = 0; i < this.layoutDGV.Columns.Count; i++)
            {
                this.layoutDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.layoutDGV.Columns[i].DefaultCellStyle.Format = "c";
            }
            for (int i = 0; i < this.feederNozzleDGV.Columns.Count; i++)
            {
                this.feederNozzleDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            //禁止编辑
            layoutDGV.ReadOnly = true;
            feederNozzleDGV.ReadOnly = true;
            summaryDGV.ReadOnly = true;

            //列和行的自动调整
            //设置行的自动自动换行，自动高度，显示出数据
            summaryDGV.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //summaryDGV.Height=
            summaryDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //列自动调整
            summaryDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //自动设置 DataGradview 的的高度无效
            //自动设置调整layoutDGV  feederNozzleDGV的高度            
            int height = GetDataGridViewHeight(layoutDGV);
            layoutDGV.Height = height;
            feederNozzleDGV.Height = height;

            Single_Evaluation.AutoScroll = true;
            //设置自动滚动的margin
            //在size事件中从新调整大小，否则不同大小的窗口显示下方会有好多空白，待修改
            if (Single_Evaluation.AutoScrollMargin.Height < summaryDGV.Location.Y +
                summaryDGV.Height + layoutDGV.Height - Single_Evaluation.Height)
            {
                Single_Evaluation.SetAutoScrollMargin(Single_Evaluation.AutoScrollMargin.Width,
                    summaryDGV.Height + layoutDGV.Height + summaryDGV.Location.Y + 60);
            }

            //每次生成后，要重新调整Single_EvaluationPanel的宽度，否则右侧会有大量空白
            //或者设置anchor，但是只宽度依靠，只有宽度自动调整大小
            //this.Single_Evaluation.Width = summaryDGV.Width;

            #endregion
            //报价单的窗体
            qotationForm = new QuotationForm(evaluationReportClass);
            //MinPrice
        }

        public int GetDataGridViewHeight(DataGridView dataGridView)
        {
            int height = dataGridView.ColumnHeadersHeight;
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                height += dataGridView.Rows[i].Height;
            }
            return height + 80;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ////要做到 单例模式，单击 当前只产生唯一的一个
            //ModefyModuleHeadCPHForm modefyModuleHeadCPHForm = ModefyModuleHeadCPHForm.Getsingle_ModefyModuleHeadCPHForm(
            //    module_Head_Cph_Structs_List, base_StatisticsDict);
            //不是单例模式，而是当前对象是否存在(!=null),则不用重新生成
            if (modefyModuleHeadCPHForm==null)
            {
                modefyModuleHeadCPHForm = new ModefyModuleHeadCPHForm(module_Head_Cph_Structs_List, base_StatisticsDict);
            }           

            modefyModuleHeadCPHForm.showForm();

            //订阅事件，保存 子窗体的数据改变到 BaseComprehensive 对象
            modefyModuleHeadCPHForm.Update_module_Head_Cph_Base += ModefyModuleHeadCPHForm_Update_module_Head_Cph_Base;
        }

        private void ModefyModuleHeadCPHForm_Update_module_Head_Cph_Base(object sender, EventArgs e)
        {
            //执行事件方法，将改动数据保存到baseComprehensive对象
            //因为改动数据都是引用类型，所有数据已将保存到对象，执行baseComprehensive保存所有数据的方法即可

            //触发 执行Single_EvaluationPanel事件
            evaluationReportClass.UpdateForSelfAfterModefy();

            //更显当前的DataGridView summary layout
            layoutDGV.DataSource = evaluationReportClass.layoutDT;
            summaryDGV.DataSource = evaluationReportClass.SummaryDT;
            //自动设置 DataGradview  调整layoutDGV 的高度            
            int height = GetDataGridViewHeight(layoutDGV);
            layoutDGV.Height = height;
            feederNozzleDGV.Height = height;
        }

        private void ViewEditQotation_Click(object sender, EventArgs e)
        {
            //if (qotationForm==null)
            //{
            //    qotationForm = new QuotationForm(evaluationReportClass);
            //}
            
            qotationForm.ShowForm();
        }

        private void CPPBtn_Click(object sender, EventArgs e)
        {
            if (qotationForm.MinPrice==0)
            {//无法生成报价单时，也无法计算
                MessageBox.Show("无法计算CPP");
                return;
            }
            CPPCalculate cPPCalculate = new CPPCalculate(qotationForm.MinPrice, 
                Convert.ToInt32(evaluationReportClass.SummaryDT.Rows[0]["CPH"].ToString()));
            cPPCalculate.Show();
            //因为传递的是值，索引最好通过扩展事件参数EventArgs e 传递数据。而且也必须放在参数中，否则，更新后的数据无法赋值过来。
            //否则需要使用showDialog()
            cPPCalculate.GetCPPEvent += UpdateCPP;
            
        }

        //更新CPP事件方法
        void UpdateCPP(object sender, CppEventArgs e)
        {
            try
            {
                //CPP赋值
                evaluationReportClass.SummaryDT.Rows[0]["CPP"] = e.CPP;

                (sender as CPPCalculate)?.Close();
            }
            catch (Exception)
            {

            }
        }
    }
}
