using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOIUse;
using NpoiOperatePicture;

namespace Exicel转换1
{
    public partial class FlexCVReport : UserControl
    {
        public FlexCVReport()
        {
            InitializeComponent();
        }

        //记录打开文件个数
        int hadOpenFileNum = 0;
        //记录打开的file，并显示
        List<string> file_List = new List<string>();
        //ExcelHelper excel处理的类
        ExcelHelper excelHelper = null;

        //承接所有打开的report 之后获取的BaseComprehensive
        List<EvaluationReportClass> baseComprehensive_list = new List<EvaluationReportClass>();

        //ExcelHelper 另一个轨道生产报告的实例
        ExcelHelper excelHelper_another = null;

        #region 打开Excel 按钮
        private void ExcelToDTbut_Click(object sender, EventArgs e)
        {
            if (System.DateTime.Now > DateTime.Parse("2023-4-10"))
            {
                return;
            }

            string file;
            if (excelHelper == null)
            {
                try
                {
                    file = OpenExcelFile();
                    //判断是否获取到文件路径
                    if (file == null)
                    {
                        return;
                    }


                    using (excelHelper = new ExcelHelper(file))
                    {
                        //打开file失败，或创建WorkBook失败
                        if (excelHelper.WorkBook == null)
                        {
                            excelHelper = null;
                            return;
                        }
                        FlexCVBaseComprehensive baseComprehensiveList = new FlexCVBaseComprehensive(excelHelper);
                        baseComprehensiveList.getBaseComprehensiveList();
                        for (int i = 0; i < baseComprehensiveList.baseComprehensive_list.Count; i++)
                        {
                            baseComprehensive_list.Add(baseComprehensiveList.baseComprehensive_list[i]);
                        }

                        hadOpenFileNum++;
                        

                        excelHelper = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
            /*
            else
            {
                if (excelHelper_another == null)
                {
                    file = OpenExcelFile();
                    //判断是否获取到文件路径
                    if (file == null)
                    {
                        return;
                    }

                    try
                    {
                        using (excelHelper_another = new ExcelHelper(file))
                        {
                            if (excelHelper_another.WorkBook == null)
                            {
                                excelHelper_another = null;
                                return;
                            }
                            //第二 另一个类中获取datatable
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("选择的两个(轨道)的生产报告还未转换！请转换后在打开其他报告，如果想选择其他转换的报告请关闭后打开！");
                    return;
                }
            }
            */



        }
        #endregion

        #region //打开Excel file 的方法
        public string OpenExcelFile()
        {
            string file = null;
            openExcelFileDialog.InitialDirectory = System.Environment.CurrentDirectory;
            openExcelFileDialog.Filter = "Excel 2007 工作表(*.xlsx)|*.xlsx|Excel 97-2003 工作表(*.xls)|*.xls";

            if (this.openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {

                if (file_List.Contains(this.openExcelFileDialog.FileName))
                {
                    MessageBox.Show(string.Format("文件{0}已经打开，请选择另外的文件", this.openExcelFileDialog.FileName));
                }
                file = this.openExcelFileDialog.FileName;//文件名
            }
            else
            {
                return null;
            }

            if (file == null)
            {
                MessageBox.Show("文件打开失败!");
                return null;
            }

            //MessageBox.Show(file);

            //DataGV.DataSource = null;
            //DataGV.DataSource = GetDataFromExcle(file);
            return file;
        }
        #endregion



        //生成评估Excel
        private void GenExlbut_Click(object sender, EventArgs e)
        {
            //未选择文件无法生成,必须excelHelper、excelHelper_another不为空时才能计算下面
            if (excelHelper == null)
            {
                MessageBox.Show("未选择生产报告，请重新选择！");
                return;
            }
            //生成需要打开第二个生产报告
            if (excelHelper_another == null)
            {
                MessageBox.Show("未选择另一个轨道的生产报告，请重新选择！");
                return;
            }

            #region 获取信息和对应的DataTable，获取结果的成功与失败，判断是否找到对应sheet和是否打开正确的Excel
            //获取位置信息从getFirstSheetDataTable中提取处理，sheet1和sheet2都有需要的地方
            //getEveryInfo();

            //第二个sheet DataTable的获取
            //验证是不是同一个轨道，Line1或Lin2，从Result获取
            string sheetName_Result = "Result";
            //主要数据从sheetName_CustomerReport获取
            string sheetName_CustomerReport = "CustomerReport_1T";
            //DataTable genDT2 = getSecondSheetDataTable(sheetName_CustomerReport, jobName);

            //第一个sheet DataTable的获取  //需要第二个中的cycletime赋值
            //DataTable genDT = getFirstSheetDataTable();
            #endregion

            //未选择文件无法生成
            //if (excelHelper == null || genDT == null)
            {
                MessageBox.Show("第一次选择的生产报告未找到对应的工作表，请确认后重新选择！");
                return;
            }
            //生成需要打开第二个生产报告
            //if (excelHelper_another == null || genDT2 == null)
            {
                MessageBox.Show("第二次选择的生产报告与第一次选择的程式名不同，或未找到对应的工作表sheet，请确认后重新选择！");
                return;
            }
            #region 

            //调用保存对话框
            string excelFileName = ComprehensiveStaticClass.SaveExcleDialogShow();

            //保存所需信息到Excel
            //ComprehensiveSaticClass.genExcelfromBaseComprehensive(excelFileName, baseComprehensive);
        }

        #region 通用方法，补齐最后的DataTable，用以生成sheet

        #endregion

        //根据获取的resultInfoTable_tuple_list重新组织summaryInfo和layoutInfo 内部DataTable
        public void regenerate_resultInfoTable_tuple_list()
        {
            //for (int i = 0; i < resultInfoTable_tuple_list.Count; i++)
            {
                //区分同一程式  对应的不同配置

                //区分同一配置，不同程式

                ////同一程式、同一配置下获取TargetConvery,区分同一程式的1、2轨信息，生成最后信息
            }
        }

        #region //根据各种信息生成第一个sheet信息的summay DataTable
        /*
    public DataTable getFirstSheetDataTable()
    {
        //生成评估的表格
        DataTable genDT = new DataTable();
        string[] genColumnString = { "Item", "Board", "Line", "Head \r\n Type", "Job name \r\n Panel Size", "Cycle time \r\n [((L1+L2)/2) Sec]", "Number of \r\n placements", "CPH", "Output \r\n Board/hour", "Output \r\n Board/Day(22H)", "CPH Rate", "Remark" };

        //评估表格的栏
        for (int i = 0; i < genColumnString.Length; i++)
        {
            genDT.Columns.Add(genColumnString[i]);
        }

        //向生成的Excel添加行
        DataRow genRow = genDT.NewRow();
        genRow[1] = boardQty;
        genRow[2] = Line;
        genRow[3] = Head_Type;
        genRow[4] = jobName;
        genRow[5] = cycleTime;
        genRow[6] = placementNumber;
        genRow[7] = cPH;
        //float.Parse(cycleTime)  float转换
        double double1 = Convert.ToDouble(boardQty) * 3600 / Convert.ToDouble(cycleTime);
        genRow[8] = string.Format("{0:0}", double1);
        genRow[9] = string.Format("{0:0}", double1 * 22);
        //42000  H24头理论的CPH，10500  H04SF理论的CPH
        //需要建一个head type:CPH的字典，进行计算；

        //CPH rate的计算，head的理论CPH*head num ++
        double CPHRate = Convert.ToDouble(cPH) / (42000 * 7 + 10500 * 1);
        genRow[10] = string.Format("{0:0.00%}", CPHRate);
        //genRow[11] = "已删减大零件";
        genRow[11] = "";

        //添加行
        genDT.Rows.Add(genRow);

        return genDT;

    }
    */
        #endregion

        #region //根据各种信息生成第二个sheet信息的DataTable
        /*
        public DataTable getSecondSheetDataTable(string sheetName_CustomerReport, string jobName)
        {
            #region //获取生成的第二个sheet对象的sheetname
            //每模组的cyccle time 在两个Excel分别获取
            //获取一轨的Excel生成对象
            getsecondsheet secondsheet = new getsecondsheet(excelHelper, sheetName_CustomerReport, jobName, conveyor, machineCount);
            bool getSecondsheetResult = secondsheet.getsecondsheetInfo();
            //获取二轨的Excel生成对象
            getsecondsheet secondsheet_another = new getsecondsheet(excelHelper_another, sheetName_CustomerReport, jobName, conveyor, machineCount);
            bool getSecondsheet_anotherResult = secondsheet_another.getsecondsheetInfo();

            if (!getSecondsheetResult)
            {
                //MessageBox.Show(string.Format("第一次选择的Excel未找到{0}表单，请重新选择！", sheetName_CustomerReport));
                //重置
                excelHelper = null;
                //未获取到sheet，则没有DataTable，返回空
                return null;
            }

            if (!getSecondsheet_anotherResult)
            {
               // MessageBox.Show(string.Format("第二次选择的Excel未找到{0}表单，请重新选择！", sheetName_CustomerReport));
                //重置
                excelHelper_another = null;
                return null;
            }

            //最终要组合的行
            //
            int secondSheetrowCount = secondsheet.SmallSecondSheetDT.Rows.Count > secondsheet.SecondSheetDT.Rows.Count ? secondsheet.SmallSecondSheetDT.Rows.Count : secondsheet.SecondSheetDT.Rows.Count;
            DataTable secondSheetDT = new DataTable();

            //初始化secondSheetDT的column
            //最终初始化的栏  i = 0; i < 15; i++)

            //设置secondSheetDT的DataTable的第一行

            secondSheetDT.Columns.Add("Prouction \n Mode");
            secondSheetDT.Columns.Add("Module \n No.");
            secondSheetDT.Columns.Add("Module");
            secondSheetDT.Columns.Add("Head Type");
            secondSheetDT.Columns.Add("Lane1 \n Cycle \n Time");
            secondSheetDT.Columns.Add("Lane2 \n Cycle \n Time");
            secondSheetDT.Columns.Add("Cycle \n Time");
            secondSheetDT.Columns.Add("Qty");
            secondSheetDT.Columns.Add("Avg");
            secondSheetDT.Columns.Add("CPH");

            secondSheetDT.Columns.Add("Feeder \n Type");
            secondSheetDT.Columns.Add("Feeder \n Qty");
            //secondSheetDT.Columns.Add("Head Type ");
            secondSheetDT.Columns.Add("Nozzle \n Type");
            secondSheetDT.Columns.Add("Nozzle \n Qty");

            /*
             *             //初始化各个栏
            secondSheetDT.Columns.Add("Prouction Mode");
            secondSheetDT.Columns.Add("Module No.");
            secondSheetDT.Columns.Add("Module");
            secondSheetDT.Columns.Add("Head Type");
            secondSheetDT.Columns.Add("Cycle Time");
            secondSheetDT.Columns.Add("Qty");

            smallsecondSheetDT.Columns.Add("Feeder");
            smallsecondSheetDT.Columns.Add("Qty Feeder");
            smallsecondSheetDT.Columns.Add("Head  Type");
            smallsecondSheetDT.Columns.Add("Nozzle");
            smallsecondSheetDT.Columns.Add("QTy  Nozzle");
             //

            //使用List，用于计算最大、平均值
            //Line1 cycletime list
            List<double> Line1CycleTimeList = new List<double>();
            //Line2 cycletime list
            List<double> Line2CycleTimeList = new List<double>();
            //cycletime list
            List<double> CycleTimeList = new List<double>();
            //Qty list
            List<int> QtyList = new List<int>();


            //循环并组合新的datatable
            for (int i = 0; i < secondSheetrowCount; i++)
            {
                DataRow dr1 = secondSheetDT.NewRow();
                secondSheetDT.Rows.Add(dr1);
                secondSheetDT.Rows[i][0] = secondsheet.SecondSheetDT.Rows[i][0];
                secondSheetDT.Rows[i][1] = secondsheet.SecondSheetDT.Rows[i][1];
                secondSheetDT.Rows[i][2] = secondsheet.SecondSheetDT.Rows[i][2];
                secondSheetDT.Rows[i][3] = secondsheet.SecondSheetDT.Rows[i][3];
                //lin1 cycyle time
                secondSheetDT.Rows[i][4] = secondsheet.SecondSheetDT.Rows[i][4];
                //secondSheetDT.Rows[i][4]  DBNull无法转换为任何其他类型
                if (!secondSheetDT.Rows[i][4].Equals(DBNull.Value))
                {
                    Line1CycleTimeList.Add(System.Convert.ToDouble(secondSheetDT.Rows[i][4]));
                }

                //lin2 cycyle time
                secondSheetDT.Rows[i][5] = secondsheet_another.SecondSheetDT.Rows[i][4];
                if (!secondSheetDT.Rows[i][5].Equals(DBNull.Value))
                {
                    Line2CycleTimeList.Add(System.Convert.ToDouble(secondSheetDT.Rows[i][5]));
                }
                //平均cycle time
                if (!secondSheetDT.Rows[i][5].Equals(DBNull.Value) && !secondSheetDT.Rows[i][4].Equals(DBNull.Value))
                {
                    secondSheetDT.Rows[i][6] = (System.Convert.ToDouble(secondSheetDT.Rows[i][4]) + System.Convert.ToDouble(secondSheetDT.Rows[i][5])) / 2;
                }


                if (!secondSheetDT.Rows[i][6].Equals(DBNull.Value))
                {

                    CycleTimeList.Add(System.Convert.ToDouble(secondSheetDT.Rows[i][6]));
                }
                //Qty
                secondSheetDT.Rows[i][7] = secondsheet.SecondSheetDT.Rows[i][5];
                if (!secondSheetDT.Rows[i][7].Equals(DBNull.Value))
                {
                    QtyList.Add(System.Convert.ToInt32(secondSheetDT.Rows[i][7]));
                }
                //Avg  问题 Cycle time/Qty
                //两位小数
                if (!secondSheetDT.Rows[i][6].Equals(DBNull.Value) && !secondSheetDT.Rows[i][7].Equals(DBNull.Value))
                {
                    secondSheetDT.Rows[i][8] = Math.Round(System.Convert.ToDouble(secondSheetDT.Rows[i][6]) / System.Convert.ToDouble(secondSheetDT.Rows[i][7]), 2);
                }

                //CPH  问题 Qty*3600/CyCle time
                //secondSheetDT.Rows[i][9] = System.Convert.ToInt32(secondSheetDT.Rows[i][7])*3600/System.Convert.ToInt32(secondSheetDT.Rows[i][6]);
                //取整
                if (!secondSheetDT.Rows[i][7].Equals(DBNull.Value))
                {
                    secondSheetDT.Rows[i][9] = string.Format("{0:0}", (System.Convert.ToInt32(secondSheetDT.Rows[i][7]) * 3600) / System.Convert.ToDouble(secondSheetDT.Rows[i][6]));
                }


                //Feeder
                secondSheetDT.Rows[i][10] = secondsheet.SmallSecondSheetDT.Rows[i][0];
                secondSheetDT.Rows[i][11] = secondsheet.SmallSecondSheetDT.Rows[i][1];
                //这一列是Head Type 列，取消
                //secondSheetDT.Rows[i][12] = secondsheet.SmallSecondSheetDT.Rows[i][2];
                secondSheetDT.Rows[i][12] = secondsheet.SmallSecondSheetDT.Rows[i][3];
                secondSheetDT.Rows[i][13] = secondsheet.SmallSecondSheetDT.Rows[i][4];
            }

            //计算最后的一行的值
            //secondSheetrowCount++;


            DataRow dr2 = secondSheetDT.NewRow();
            secondSheetDT.Rows.Add(dr2);
            secondSheetDT.Rows[secondSheetrowCount][0] = "Total";
            secondSheetDT.Rows[secondSheetrowCount][4] = Line1CycleTimeList.Max();
            secondSheetDT.Rows[secondSheetrowCount][5] = Line2CycleTimeList.Max();
            secondSheetDT.Rows[secondSheetrowCount][6] = CycleTimeList.Max();
            secondSheetDT.Rows[secondSheetrowCount][7] = QtyList.Sum(); ;
            secondSheetDT.Rows[secondSheetrowCount][8] = string.Format("{0:0.00}", CycleTimeList.Max() * machineCount / QtyList.Sum());
            secondSheetDT.Rows[secondSheetrowCount][9] = string.Format("{0:0}", QtyList.Sum() * 3600 / CycleTimeList.Max());

            //为第一个sheet的cycletime赋值
            cycleTime = string.Format("{0:0.00}", CycleTimeList.Max());

            //验证获取到的DataTable
            return secondSheetDT;
            #endregion
        }
        */
        #endregion


        #endregion
        private void importFlexCVCReport_Load(object sender, EventArgs e)
        {

        }
    }
}
