using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
    public partial class Form1 : Form
    {
        //ExcelHelper excel处理的类
        ExcelHelper excelHelper = null;

        //ExcelHelper 另一个轨道生产报告的实例
        ExcelHelper excelHelper_another = null;

        //dt：Excel表格
        public DataTable dt = null;
        //三个位置点
        public int[] Pannel_Point;
        public int[] Calculation_Point;
        public int[] Machineconf_Point;

        //  JOB总数,若果有多个程式也记录下来
        int[][] JOBs;
        // pattern 总数
        int[][] calculation_Pattern_POints;
        int[][] machineconf_Pattern_POints;

        // 模组数
        int machineCount;
        // 机器种类
        string machineKind;
        // 所有的Module Type
        string[] allModuleTypeString;
        //所有的HEAD Type
        string[] allHeadTypeString;

        //字典，存放<模组类型:模组个数>
        Dictionary<string, int> module_Statistics;
        //模组种类，即module_Statistics的key
        ////M3、M6三代、二代\一代的使用
        //string moduleKind3_III = "M3III";
        //string moduleKind6_III = "M6III";
        //string moduleKind3_II = "M3II";
        //string moduleKind6_II = "M6II";
        //string moduleKind3_I = "M3I";
        //string moduleKind6_I = "M6I";

        //字典，存放<HeadType:头个数>
        Dictionary<string, int> Head_Statistics;

        //中间图片位置信息
        string Line_short;
        //图片信息
        List<PicturesInfo> pictureInfoList = null;

        //---生成Excel各个栏的数据
        int boardQty;
        string Line;
        string Head_Type;
        string jobName;
        string cycleTime;
        string placementNumber;
        string cPH;
        string conveyor;
        string TargetConveyor;

        public Form1()
        {
            InitializeComponent();
        }

        #region 打开Excel对话框
        private void ExcelToDTbut_Click(object sender, EventArgs e)
        {
            if (System.DateTime.Now > DateTime.Parse("2021-4-10"))
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
                        //ExcelToDataTabl获取sheet到DataTable，第一行是否作为DataTable的列名
                        //此处获取dt，是由于其他值的计算需要先获取Result Sheet
                        dt = excelHelper.ExcelToDataTable("Result", false);
                        //dt = excelHelper.ExcelToDataTable(null, true);

                        //如果打开没有Result的表格，打开其他非报告的表格，则返回DataTable null，重置
                        if (dt == null)
                        {
                            MessageBox.Show(string.Format("未找到{0}表单，请重新选择！", "Result"));
                            excelHelper = null;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
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

        #region 获取result对应的dt，从Excel到DataTable
        // 获取result对应的dt，从Excel到DataTable

        # endregion

        #region 获取result下的位置信息

        //获取区间三个位置点    "Pannel information"  "Calculation results"  "Machine configration'  的 位置
        public void getPannelCalculationMachineconf()
        {
            //获取三个位置点
            string findString1 = "Panel information";
            string findString2 = "Calculation results";
            string findString3 = "Machine configration";
            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt);
            int[][] Pannel_Points = sectionInfo.GetInfo_Between(findString1);
            int[][] Calculation_Points = sectionInfo.GetInfo_Between(findString2);
            int[][] Machineconf_Points = sectionInfo.GetInfo_Between(findString3);

            Pannel_Point = Pannel_Points[0];
            Calculation_Point = Calculation_Points[0];
            Machineconf_Point = Machineconf_Points[0];
        }

        //获取所有JOB位置
        public void getJOBs()
        {
            //获取的JOB字符
            string findString1 = "JOB";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Pannel_Point, Calculation_Point);
            JOBs = sectionInfo.GetInfo_Between(findString1);

        }

        //获取Calculation results下的所有 Pattern 位置
        //Pattern1\Pattern2\Pattern3  Pattern格式
        public void getCalculationPatterns()
        {
            //获取的Pattern 字符，不是完全匹配
            string findString1 = "Pattern";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Calculation_Point, Machineconf_Point);

            calculation_Pattern_POints = sectionInfo.GetInfo_Between(findString1, false);

        }


        //获取Machine configration下的所有 Pattern 位置
        //Pattern1\Pattern2\Pattern3  Pattern格式
        public void getMachineconfPatterns()
        {
            //获取的Pattern 字符，不是完全匹配
            string findString1 = "Pattern";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Machineconf_Point);

            machineconf_Pattern_POints = sectionInfo.GetInfo_Between(findString1, false);

        }
        #endregion

        #region 一些通用方法

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，从overThisPoints位置之上，根据指定的字符串，获取此字符串下一行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string getSpecifyString_NextRow(string findString, int[] underThisPoints, int[] overThisPoints)
        {
            string[] destString = getSpecifyString_NextRow(findString, underThisPoints, overThisPoints, 1);
            return destString[0];
        }

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，从overThisPoints位置之上，根据指定的字符串，获取此字符串下一行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string getSpecifyString_NextRow(string findString, int[] underThisPoints, int RowX, int ColumnY)
        {
            string[] destString = getSpecifyString_NextRow(findString, underThisPoints, new int[] { underThisPoints[0] + RowX, underThisPoints[1] + ColumnY }, 1);
            return destString[0];
        }

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，从overThisPoints位置之上，根据指定的字符串，获取此字符串下findN行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string[] getSpecifyString_NextRow(string findString, int[] underThisPoints, int[] overThisPoints, int findN)
        {
            int[][] findString_Points;
            //findString_points的位置
            getDataTablePointsInfo findStringPoint;
            //判断有无结束位置
            if (overThisPoints == null)
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(dt, underThisPoints);
            }
            else
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(dt, underThisPoints, overThisPoints);
            }
            findString_Points = findStringPoint.GetInfo_Between(findString);

            if (findString_Points == null)
            {
                //没有找到字符串
                return null;
            }

            //获取findString_points位置下面findN行的内容
            List<string> theDestString = new List<string>();
            for (int i = 0; i < findN; i++)
            {
                int[] lastFindPoints = new int[] { findString_Points[0][0] + (i + 1), findString_Points[0][1] };
                theDestString.Add(findStringPoint.GetInfo_Between(lastFindPoints));
            }

            return theDestString.ToArray();
        }

        /// <summary>
        /// 通用方法  重载，从underThisPoints位置开始，无结束位置，根据指定的字符串，获取此字符串下一行的内容
        /// </summary>
        /// <returns></returns>
        public string getSpecifyString_NextRow(string findString, int[] underThisPoints)
        {
            string[] destString = getSpecifyString_NextRow(findString, underThisPoints, 1);
            return destString[0];
        }

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，无结束位置，根据指定的字符串，获取此字符串下findN行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string[] getSpecifyString_NextRow(string findString, int[] underThisPoints, int findN)
        {
            return getSpecifyString_NextRow(findString, underThisPoints, null, findN);

        }

        /// <summary>
        /// 通用方法，通过开始的字符串获取一个以字符串开始位置的一段DataTable区域
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public DataTable getCell_DT(string startString, int rowX, int columnY)
        {

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, startString);

            GetCellsDefined getCells = new GetCellsDefined();

            getCells.GetCellArea(dt, sectionInfo.StartPoint, rowX, columnY);


            return getCells.DT;
        }


        /// <summary>
        /// 通用方法，通过开始字符串获取一个以字符串开始位置，此字符串右侧的的数据
        /// </summary>

        public string getSpecifyString_NextColumn(string findString, int[] underThisPoints, int[] overThisPoints = null)
        {
            //findString对应的位置
            int[][] findString_Points;
            //findString_points的位置的对象
            getDataTablePointsInfo findStringPoint;
            //判断有无结束位置
            if (overThisPoints == null)
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(dt, underThisPoints);
            }
            else
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(dt, underThisPoints, overThisPoints);
            }
            findString_Points = findStringPoint.GetInfo_Between(findString);

            if (findString_Points == null)
            {
                //没有找到字符串
                return null;
            }
            int[] lastFindPoints = new int[] { findString_Points[0][0], findString_Points[0][1] + 1 };
            //theDestString.Add(findStringPoint.GetInfo_Between(lastFindPoints));
            return findStringPoint.GetInfo_Between(lastFindPoints);
        }
        #endregion

        #region //获取 生成报告的sheet1各个信息和位置点
        private void getEveryInfo()
        {
            getPannelCalculationMachineconf();
            getJOBs();
            getCalculationPatterns();
            getMachineconfPatterns();


            /*
            MessageBox.Show("Pannel_Point：" + Pannel_Point[0] + Pannel_Point[1]);
            MessageBox.Show("Calculation_Point：" + Calculation_Point[0] + Calculation_Point[1]);
            MessageBox.Show("Machineconf_Point：" + Machineconf_Point[0] + Machineconf_Point[1]);
            */


            //读取JOBS
            //for (int i = 0; i < JOBs.Rank; i++)
            for (int i = 0; i < 1; i++)
            {
                //MessageBox.Show("JOBs：" + JOBs[i][0] + JOBs[i][1]);

                //获取board数
                ///   1、获取JOB区域,JOB区域 3行 17列
                ///   2、从JOB区域获取获取Board qty位置，获取数量
                ///  
                string findString_BoardQty = "Board qty";
                string BoardQtyString = getSpecifyString_NextRow(findString_BoardQty, JOBs[i], 3, 17);

                boardQty = Convert.ToInt32(BoardQtyString);

                string findString_jobName = "Name";
                jobName = getSpecifyString_NextRow(findString_jobName, JOBs[i], 3, 17);

                //cycletime是line1和 line2的两个时间的平均，但是是在两个报告，两个Excel里面
                //string findString_cycleTime = "Cycle time";
                //string cycleTime_line1 = getSpecifyString_NextRow(findString_cycleTime, JOBs[i], 3, 17);
                //cycleTime = cycleTime_line1;

                string findString_partQty = "Part qty";
                placementNumber = getSpecifyString_NextRow(findString_partQty, JOBs[i], 3, 17);
            }


            //读取calculation Patterns
            //for (int i = 0; i < calculation_Pattern_POints.Rank; i++)
            for (int i = 0; i < 1; i++)
            {
                //MessageBox.Show("calculation_Pattern_POints：" + calculation_Pattern_POints[i][0] + calculation_Pattern_POints[i][1]);
            }

            //读取machineconf Patterns,及其下的 machine count
            //for (int i = 0; i < machineconf_Pattern_POints.Rank; i++),若真循环会不断覆盖值，多个job时需要做到动态添加，或者没获取一个信息则处理一个信心。
            //.Rank获取的数据也不对，.Rank获取的是维数，不是第一维的个数
            for (int i = 0; i < 1; i++)
            {
                // MessageBox.Show("machineconf_Pattern_POints：" + machineconf_Pattern_POints[i][0] + machineconf_Pattern_POints[i][1]);
                //获取机器的count
                string findString_machineCount = "Machine count";
                string machineCountString = getSpecifyString_NextRow(findString_machineCount, machineconf_Pattern_POints[i]);
                machineCount = Convert.ToInt32(machineCountString);
                // MessageBox.Show("machineCount：" + machineCount);

                //获取机器种类
                string findString_machineKind = "Machine kind";
                string machineKindString = getSpecifyString_NextRow(findString_machineKind, machineconf_Pattern_POints[i]);
                machineKind = machineKindString;


                //获取组合line的信息
                //////获取所有的Module Type,
                string findString_AllModuleType = "Module Type";
                allModuleTypeString = getSpecifyString_NextRow(findString_AllModuleType, machineconf_Pattern_POints[i], machineCount);

                //创建统计模组信息的字典
                module_Statistics = new Dictionary<string, int>();
                //初始化，仅记录M3III，M6III，初始为0

                #region 使用先添加key再统计信息的方法，不灵活
                //module_Statistics.Add(moduleKind3_III, 0);
                //module_Statistics.Add(moduleKind6_III, 0);

                //for (int j = 0; j < allModuleTypeString.Length; j++)
                //{
                //    if (allModuleTypeString[j] == moduleKind3_III)
                //    {
                //        module_Statistics[moduleKind3_III] += 1;
                //    }
                //    else if (allModuleTypeString[j] == moduleKind6_III)
                //    {
                //        module_Statistics[moduleKind6_III] += 1;
                //    }

                //}

                ////拼接Line字符串
                ////计算base信息
                //int baseCount_4M = ((module_Statistics[moduleKind3_III] + module_Statistics[moduleKind6_III] * 2) / 4);
                //int baseCount_2M;
                //if (((module_Statistics[moduleKind3_III] + module_Statistics[moduleKind6_III] * 2) % 4) != 0)
                //{
                //    baseCount_2M = 1;
                //}
                //else
                //{
                //    baseCount_2M = 0;
                //}

                ////"/r/n" 回车换行符
                //Line = "";
                //Line_short = "";
                //Line_short = machineKind + "-(" + moduleKind3_III + "*" + module_Statistics[moduleKind3_III].ToString() + "+" + moduleKind6_III + "*" + module_Statistics[moduleKind6_III] + ")";
                //Line = machineKind + "-(" + moduleKind3_III + "*" + module_Statistics[moduleKind3_III].ToString() + "+" + moduleKind6_III + "*" + module_Statistics[moduleKind6_III] + ")" + "\n" + "4M BASE III*" + baseCount_4M.ToString() + "+2M BASE III*" + baseCount_2M.ToString();
                ////MessageBox.Show(Line);
                #endregion

                #region 重新统计模组信息
                for (int j = 0; j < allModuleTypeString.Length; j++)
                {
                    if (module_Statistics.ContainsKey(allModuleTypeString[j]))
                    {
                        module_Statistics[allModuleTypeString[j]]++;
                    }
                    else
                    {
                        module_Statistics.Add(allModuleTypeString[j], 1);
                    }
                }
                #endregion

                //拼接Line字符串

                //重新计算base信息
                //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
                //base尽可能少的分配
                int Count_M = 0;
                foreach (var moduleKind in module_Statistics)
                {
                    if (moduleKind.Key.Substring(0, 2) == "M3")
                    {
                        Count_M += moduleKind.Value;
                    }
                    if (moduleKind.Key.Substring(0, 2) == "M6")
                    {
                        Count_M += moduleKind.Value * 2;
                    }
                }
                int baseCount_4M = (Count_M / 4);
                int baseCount_2M = 0;
                if ((Count_M % 4) != 0)
                {
                    baseCount_2M = 1;
                }

                //"/r/n" 回车换行符
                Line = "";
                Line_short = "";
                foreach (var item in module_Statistics)
                {
                    Line_short += item.Key + "*" + item.Value.ToString() + "+";
                }
                Line_short = machineKind + "-(" + Line_short.Substring(0, Line_short.Length - 1) + ")";
                //判断4Mhe 2M base的数量，进行拼接
                if (baseCount_2M == 0 && baseCount_4M != 0)
                {
                    Line = Line_short + "\n" + "4M BASE III*" + baseCount_4M.ToString();
                }
                if (baseCount_2M != 0 && baseCount_4M != 0)
                {
                    Line = Line_short + "\n" + "4M BASE III*" + baseCount_4M.ToString() + "+2M BASE III*" + baseCount_2M.ToString();
                }
                if (baseCount_2M != 0 && baseCount_4M == 0)
                {
                    Line = Line_short + "\n" + "2M BASE III*" + baseCount_2M.ToString();
                }

                //MessageBox.Show(Line);

                //获取组合"Head Type"的信息
                //////获取所有的"Head Type",
                string findString_AllHeadType = "Head Type";
                allHeadTypeString = getSpecifyString_NextRow(findString_AllHeadType, machineconf_Pattern_POints[i], machineCount);

                //统计头的信息 字典
                Head_Statistics = new Dictionary<string, int>();

                for (int j = 0; j < allHeadTypeString.Length; j++)
                {
                    // 只有两种情况，包含key和不包含key
                    if (Head_Statistics.ContainsKey(allHeadTypeString[j]))
                    {
                        Head_Statistics[allHeadTypeString[j]] += 1;
                    }
                    else //if (!Head_Statistics.ContainsKey(allHeadTypeString[j]))
                    {
                        Head_Statistics.Add(allHeadTypeString[j], 1);
                    }

                }

                //拼接Head Type
                Head_Type = "";
                foreach (var item in Head_Statistics)
                {
                    Head_Type += item.Key + "*" + item.Value + "+";
                }
                Head_Type = Head_Type.Substring(0, Head_Type.Length - 1);
                // MessageBox.Show(Head_Type);

                #region
                //NPOI导出生成图片，不用判断格式，jpg和png均可生成。
                //获取图片
                pictureInfoList = new List<PicturesInfo>();
                ISheet sheet = excelHelper.ExcelToIsheet("Result");
                int maxColumn = ExcelHelper.sheetColumns(sheet);
                //仅处理只有一个程式的情况，machineconf_Pattern_POints，到文档结束。
                //pictureInfoList = NpoiExtend.GetAllPictureInfos(sheet, machineconf_Pattern_POints[i][0], sheet.LastRowNum, machineconf_Pattern_POints[i][1], maxColumn);

                #endregion


                //获取CPH
                string findString_CPH = "CPH";
                cPH = getSpecifyString_NextRow(findString_CPH, machineconf_Pattern_POints[i]);

                //获取Conveyor,Dual
                string findString_Conveyor = "Conveyor";
                conveyor = getSpecifyString_NextColumn(findString_Conveyor, machineconf_Pattern_POints[i]);

                //获取
                string findString_TargetConveyor = "TargetConveyor";
                TargetConveyor = getSpecifyString_NextColumn(findString_TargetConveyor, machineconf_Pattern_POints[i]);

            }
        }
        #endregion




        //生成DataTable
        //生成评估sheet1
        //sheet2未验证jobName job名字
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
            getEveryInfo();

            //第二个sheet DataTable的获取
            //验证是不是同一个轨道，Line1或Lin2，从Result获取
            string sheetName_Result = "Result";
            //主要数据从sheetName_CustomerReport获取
            string sheetName_CustomerReport = "CustomerReport_1T";
            DataTable genDT2 = getSecondSheetDataTable(sheetName_CustomerReport, jobName);

            //第一个sheet DataTable的获取  //需要第二个中的cycletime赋值
            DataTable genDT = getFirstSheetDataTable();
            #endregion

            //未选择文件无法生成
            if (excelHelper == null || genDT == null)
            {
                MessageBox.Show("第一次选择的生产报告未找到对应的工作表，请确认后重新选择！");
                return;
            }
            //生成需要打开第二个生产报告
            if (excelHelper_another == null || genDT2 == null)
            {
                MessageBox.Show("第二次选择的生产报告与第一次选择的程式名不同，或未找到对应的工作表sheet，请确认后重新选择！");
                return;
            }


            #region 保存对话框的实现
            //保存的路径
            //string fileName = @"D:\Test\110.xlsx";
            string fileName = null;
            saveExcelFileDialog.Filter = "Excel 2007 工作表(*.xlsx)|*.xlsx|Excel 97-2003 工作表(*.xls)|*.xls";
            saveExcelFileDialog.RestoreDirectory = true;
            if (saveExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    fileName = saveExcelFileDialog.FileName;
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

            }
            else
            {
                return;
            }

            //使用指定的fileName 创建Iworkbook
            if (fileName == null)
            {
                MessageBox.Show("发生了意外！");
                return;
            }
            #endregion


            //创建ExcelHelper，用以存放转换后的sheet1和sheet2
            ExcelHelper excelHelper1 = new ExcelHelper(fileName);

            #region 第一个sheet的获取和生成
            //第一个sheet的标题，以字符串形式插入
            string oneTitle = "FUJI Line Throughput Report " + machineKind;

            string sheet1 = "sheet1";

            //插入sheet对应表格的datatable
            DataTableToResultSheet1(excelHelper1, genDT, sheet1, true, new int[2] { 3, 1 });
            //插入标题
            StringInsertExcel(excelHelper1, oneTitle, sheet1, new int[2] { 1, 1 }, IndexedColors.BrightGreen.Index);

            //插入图片位置信息
            StringInsertExcel(excelHelper1, "", sheet1, new int[2] { 2, 1 }, IndexedColors.LightGreen.Index);
            //Line_short 需要以画布的形式写入，而不是单元格内的文字;
            //插入图片
            genFirstSheetPicture(excelHelper1, sheet1, new int[2] { 2, genDT.Columns.Count / 2 - 2 }, module_Statistics);
            #endregion

            #region 生成第二个sheet
            //第二个sheet的标题，以字符串形式插入
            string secondTitle = jobName;
            //第二个sheet的建议
            string ProposalString = "Proposal \n Layout";

            //获取生成的第二个sheet对象在Excel对应的sheetname
            string sheet2 = "sheet2";

            //插入sheet对应表格的datatable
            DataTableToResultSheet2(excelHelper1, genDT2, sheet2, true, new int[2] { 3, 2 });
            //插入标题
            StringInsertExcel(excelHelper1, secondTitle, sheet2, new int[2] { 2, 2 }, IndexedColors.BrightGreen.Index);

            //插入建议
            StringInsertExcel(excelHelper1, ProposalString, sheet2, new int[2] { 2, 1 }, IndexedColors.BrightGreen.Index);
            #endregion


            //获取对象excelHelper1的文件流，即使用filename创建的文件流，判断是否有文件，没有文件的fa为null
            FileStream fs = excelHelper1.FS;
            if (fs == null)
            {
                using (fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
                    fs.Close();
                }

            }
            else
            {
                excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
                fs.Close();
            }


            MessageBox.Show("生成Excel成功！");
            //生成成功，重置两个实例
            excelHelper = null;
            excelHelper_another = null;
        }

        /// <summary>
        /// 生成图片
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint"></param>
        public void genFirstSheetPicture(ExcelHelper excelHelper1, string sheetName, int[] insertPoint, Dictionary<string, int> module_Statistics)
        {
            if (pictureInfoList == null)
            {
                return;
            }

            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;

            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];

            ISheet sheet1 = excelHelper1.ExcelToIsheet(sheetName);

            string img_M3 = @".\img\M3.png";
            string img_M6 = @".\img\M6.png";

            //获取两个文件的img
            var imgM3 = Image.FromFile(img_M3);
            var imgM6 = Image.FromFile(img_M6);

            //获取要生成的图片的宽高
            //生成最后图片的外边框留白,jiange间隔
            int jiange = 5;
            //int finalWidth = imgM3.Width * module_Statistics[moduleKind3_III]/2 + imgM6.Width*module_Statistics[moduleKind6_III] + jiange*2;

            //统计所需要拼接的图片数量
            int M3Image_count = 0;
            int M6Image_count = 0;
            foreach (var item in module_Statistics)
            {
                if (item.Key.Substring(0, 2) == "M3")
                {
                    M3Image_count += item.Value / 2;
                }
                if (item.Key.Substring(0, 2) == "M6")
                {
                    M6Image_count += item.Value;
                }
            }
            int finalWidth = imgM3.Width * M3Image_count + imgM6.Width * M6Image_count + jiange * 2;
            int finalHeight = imgM3.Height + jiange * 2;

            Bitmap finalImg = new Bitmap(finalWidth, finalHeight);
            Graphics graph = Graphics.FromImage(finalImg);
            //graph.Clear(SystemColors.AppWorkspace);
            graph.Clear(Color.Empty);

            //循环模组并获取img画图
            //获取的位置
            int pointX = jiange;
            int pointY = jiange;

            for (int i = 0; i < allModuleTypeString.Length; i++)
            {
                //M3模组图像
                if (allModuleTypeString[i].Substring(0, 2) == "M3" && allModuleTypeString[i + 1].Substring(0, 2) == "M3")
                {
                    graph.DrawImage(imgM3, pointX, pointY);
                    pointX += imgM3.Width;
                    i++;
                }
                else if (allModuleTypeString[i].Substring(0, 2) == "M6")//M6模组图像
                {
                    graph.DrawImage(imgM6, pointX, pointY);
                    pointX += imgM6.Width;
                }
            }
            graph.Dispose();

            //将bitmap转换为byte[]
            byte[] PictureData;
            using (MemoryStream stream = new MemoryStream())
            {
                finalImg.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                PictureData = new byte[stream.Length];
                stream.Seek(0, SeekOrigin.Begin);
                stream.Read(PictureData, 0, Convert.ToInt32(stream.Length));
            }


            //插入的位置
            int startRow = count;
            int startCol = insertPoint[1];
            int endRow = startRow + 1;
            int endCol = startCol + 2;
            //根据模组数量长度，变更结束单元格位置
            if (M3Image_count + M6Image_count > 7)
            {
                endCol = startCol + 3;
            }
            if (M3Image_count + M6Image_count > 14)
            {
                endCol = startCol + 4;
            }

            //偏移依旧不起作用
            excelHelper1.pictureDataToSheet(sheet1, PictureData, 50, 10, 50, 10, startRow, startCol, endRow, endCol);
        }
        /*
    public void genFirstSheetPicture(ExcelHelper excelHelper1, string sheetName,int[] insertPoint)
    {
        if (pictureInfoList==null)
        {
            return;
        }

        //生成报告的excelHelper1 excelHelper类
        IWorkbook workbook = excelHelper1.WorkBook;

        //从指定的位置开始插入DataTable数据
        int count = insertPoint[0];

        ISheet sheet1 = excelHelper1.ExcelToIsheet(sheetName);


        int startRow= count;
        int startCol= insertPoint[1];
        int endRow = startRow + 1;
        int endCol = startCol + 1;

        //单元格内的位置,
        int dx1 = 0;
        int dy1 = 0;
        int dx2 = 500;
        int dy2 = 0;

        //获取单元格
        //IRow row = sheet1.GetRow(startRow);
        //ICell cell = sheet1.GetRow(startRow).GetCell(startCol);
        //pictureInfoList需要排序
        foreach (PicturesInfo pictureInfo in pictureInfoList)
        {

            if (pictureInfoList.IndexOf(pictureInfo) == 0)
            {                  
            }
            else
            {
                //图片开始偏移量
                dx1=dx1 + 500;

                if (dx1>=1023)
                {
                    dx1 = dx1 - 1023;
                    startCol++;
                    endCol++;
                }

                //图片结束偏移量
                //当宽高小于偏移量时，移动到下一单元格
                if (1023 <= dx1 + 500)
                {
                    dx2 = (dx1 + 500) - 1023;
                    endCol = endCol + 1;
                }
                else
                {
                    //下一个图片的开始，是上一个图片的结束，总宽1023-dy2
                    dx2 = dx1 + 500;
                }           
            }
             excelHelper1.pictureDataToSheet(sheet1, pictureInfo.PictureData,dx1,dy1,dx2,dy2, startRow, startCol, endRow, endCol);

        }
    }
    */

        /// <summary>
        /// 设置单元格背景
        /// </summary>已经生成过单元格，再次设置背景无效.已验证。需要在第一次创建cell时填充颜色才可以
        /// <returns></returns>
        /// 


        #region //根据各种信息生成第一个sheet信息的DataTable
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

            //像生成的Excel添加行
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
            double CPHRate = Convert.ToDouble(cPH) / (42000 * 7 + 10500 * 1);
            genRow[10] = string.Format("{0:0.00%}", CPHRate);
            //genRow[11] = "已删减大零件";
            genRow[11] = "";

            //添加行
            genDT.Rows.Add(genRow);

            return genDT;



        }
        #endregion

        #region //根据各种信息生成第二个sheet信息的DataTable
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
             */

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
        #endregion



        //将string插入Excel指定位置
        public void StringInsertExcel(ExcelHelper excelHelper1, string insertString, string sheetName, int[] insertPoint, short colorShort)
        {

            IWorkbook workbook = excelHelper1.WorkBook;
            //FileStream fs = excelHelper1.FS;

            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];
            ISheet sheet = excelHelper1.ExcelToIsheet(sheetName);

            //sheet由excelHelper获取
            try
            {
                IRow row;
                if (sheet.GetRow(count) == null)
                {
                    row = sheet.CreateRow(count);
                }
                else
                {
                    row = sheet.GetRow(count);
                }

                //row.HeightInPoints = 15;
                row.CreateCell(insertPoint[1]).SetCellValue(insertString);


                #region 样式
                ICellStyle cellStyle = workbook.CreateCellStyle();
                //垂直居中
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //水平对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                cellStyle.WrapText = true;

                //设置背景
                cellStyle.FillBackgroundColor = IndexedColors.White.Index;
                cellStyle.FillForegroundColor = colorShort;
                cellStyle.FillPattern = FillPattern.SolidForeground;

                //设置边框在，在合并的单元格中，边框未应用到整个合并单元格
                //cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;

                //应用到单元格样式
                row.GetCell(insertPoint[1]).CellStyle = cellStyle;
                #endregion


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelHelper1">要生成的Excel对应的ExcelHelper类对象</param>
        /// <param name="dt">生成sheet的DataTable</param>
        /// ///////////////////<param name="fileName">生成的文件，在判断fs没有的情况下，创建fs文件流</param>
        /// <param name="sheetNamet">将dt插入到指定sheet</param>
        /// <param name="isColumnWritten">是否吧列名作为单元格插入</param>
        /// <param name="insertPoint">插入的位置</param>
        public void DataTableToResultSheet1(ExcelHelper excelHelper1, DataTable dt, string sheetName, bool isColumnWritten, int[] insertPoint)
        {


            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;
            //fs 作为向现有文件生成表格还是新建文件生成表格
            //FileStream fs = excelHelper1.FS;



            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];

            ISheet sheet = excelHelper1.ExcelToIsheet(sheetName);

            #region sheet由excelHelper获取
            //if (workbook != null)
            //{
            //    //有则获取
            //    if (workbook.GetSheetIndex(sheetNamet) >= 0)
            //    {
            //        sheet = workbook.GetSheet(sheetNamet);
            //    }
            //    else
            //    {
            //        sheet = workbook.CreateSheet(sheetNamet);
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("失败");
            //    return;
            //}

            #endregion
            try
            {
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Columns[j].ColumnName);
                    }
                    count++;
                }


                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                    ++count;
                }


                #region sheet1的样式
                //设置列宽
                int commonWidth = 10;
                sheet.SetColumnWidth(0, 1 * 256);
                sheet.SetColumnWidth(1, commonWidth * 256);
                sheet.SetColumnWidth(2, commonWidth * 256);
                sheet.SetColumnWidth(3, System.Convert.ToInt32(System.Convert.ToDouble(commonWidth * 256) * 4.5));
                sheet.SetColumnWidth(4, commonWidth * 2 * 256);
                sheet.SetColumnWidth(5, commonWidth * 2 * 256);
                sheet.SetColumnWidth(6, commonWidth * 256);
                sheet.SetColumnWidth(7, commonWidth * 256);
                sheet.SetColumnWidth(8, commonWidth * 256);
                sheet.SetColumnWidth(9, commonWidth * 256);
                sheet.SetColumnWidth(10, commonWidth * 256);
                sheet.SetColumnWidth(11, commonWidth * 256);
                sheet.SetColumnWidth(12, commonWidth * 256);



                //合并单元格
                sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, (dt.Columns.Count - 1) + 1));//2.0使用 2.0以下为Region
                sheet.AddMergedRegion(new CellRangeAddress(2, 2, 1, (dt.Columns.Count - 1) + 1));//2.0使用 2.0以下为Region



                //设置行高
                IRow row1,
                    row2,
                    row3,
                    row4;
                if (sheet.GetRow(1) == null)
                {
                    row1 = sheet.CreateRow(1);
                }
                else
                {
                    row1 = sheet.GetRow(1);
                }

                if (sheet.GetRow(2) == null)
                {
                    row2 = sheet.CreateRow(2);
                }
                else
                {
                    row2 = sheet.GetRow(2);
                }

                if (sheet.GetRow(3) == null)
                {
                    row3 = sheet.CreateRow(3);
                }
                else
                {
                    row3 = sheet.GetRow(3);
                }

                if (sheet.GetRow(4) == null)
                {
                    row4 = sheet.CreateRow(4);
                }
                else
                {
                    row4 = sheet.GetRow(4);
                }



                row1.HeightInPoints = 50;
                row2.HeightInPoints = 80;
                row3.HeightInPoints = 30;
                row4.HeightInPoints = 50;

                //设置整列单元格样式
                //获取列数量
                //int column_count = ExcelHelper.sheetColumns(sheet);
                //for (int i = 0; i < column_count; i++)
                //{
                //    //设置SetDefaultColumnStyle后，仅到新创建cell时，会应用这个默认设置
                //    //sheet.SetDefaultColumnStyle(i, cellStyle);
                //}
                ICellStyle cellStyle;
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    if (sheet.GetRow(i) != null)
                    {
                        IRow row = sheet.GetRow(i);

                        for (int j = 0; j <= row.LastCellNum; j++)
                        {
                            //int j = row.FirstCellNum,这个复制在第一次循环时提示cell number必须>=0，FirstCellNum怎么会<0？
                            if (row.GetCell(j) == null)
                            {
                                continue;
                            }

                            //设置单元格样式,必须每一个cell对饮一个样式，否则无法设置背景

                            cellStyle = workbook.CreateCellStyle();
                            //垂直居中
                            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                            //水平对齐
                            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                            //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                            cellStyle.WrapText = true;

                            //设置cellstyle背景
                            cellStyle.FillBackgroundColor = IndexedColors.Black.Index;
                            cellStyle.FillForegroundColor = IndexedColors.LightGreen.Index;

                            //设置边框
                            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

                            if (i == sheet.LastRowNum)
                            {
                                cellStyle.FillForegroundColor = IndexedColors.White.Index;
                                if (j == row.LastCellNum)
                                {
                                    //创建字体和设置字号
                                    IFont font = workbook.CreateFont();
                                    font.FontHeightInPoints = 6;
                                    cellStyle.SetFont(font);
                                }
                            }

                            //填充模式
                            cellStyle.FillPattern = FillPattern.SolidForeground;

                            row.GetCell(j).CellStyle = cellStyle;
                            //cellStyle.Dispose();
                        }
                    }
                }
                #endregion

                #region 写入流应该在方法外
                //if (fs == null)
                //{
                //    fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //}
                //workbook.Write(fs); //写入到excel文件，并关闭流
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        public void DataTableToResultSheet2(ExcelHelper excelHelper1, DataTable dt, string sheetNamet, bool isColumnWritten, int[] insertPoint)
        {
            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;

            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];

            ISheet sheet = excelHelper1.ExcelToIsheet(sheetNamet);

            try
            {
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Columns[j].ColumnName);
                    }
                    count++;
                }


                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                    ++count;
                }

                setSecondSheetStyle(sheet, dt.Columns.Count - 1);

                #region 写入流应该在方法外
                //if (fs == null)
                //{
                //    fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //}
                //workbook.Write(fs); //写入到excel文件，并关闭流
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        public void setSecondSheetStyle(ISheet sheet, int mergeLastCount)
        {
            //应使用静态方法，可以获取不规则行列的sheet的最大与最小列
            int sheetColumns = ExcelHelper.sheetColumns(sheet);
            #region sheet2的样式
            //设置列宽
            int commonWidth = 8;
            sheet.SetColumnWidth(0, 1 * 256);
            for (int i = 1; i < sheetColumns; i++)
            {
                if (i == sheetColumns - 2)
                {
                    sheet.SetColumnWidth(i, System.Convert.ToInt32(commonWidth * 1.8) * 256);
                    continue;
                }
                if (i == sheetColumns - 4)
                {
                    sheet.SetColumnWidth(i, System.Convert.ToInt32(commonWidth * 1.8) * 256);
                    continue;
                }
                sheet.SetColumnWidth(i, commonWidth * 256);
            }

            //设置行高
            //是否是有内容的第一行，有内容的第一行上面和左边合并
            bool isFirstDataRow = true;
            //workbook，创建cellstyle
            IWorkbook workbook = sheet.Workbook;
            ICellStyle cellStyle;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                if (sheet.GetRow(i) != null)
                {
                    //设置第一次数据行的行高和和合并，在cell循环之外。
                    if (isFirstDataRow)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(i - 1, i - 1, 2, 2 + mergeLastCount));
                        sheet.AddMergedRegion(new CellRangeAddress(i - 1, sheet.LastRowNum, 1, 1));
                        sheet.GetRow(i).HeightInPoints = 50;
                        sheet.CreateRow(i - 1).HeightInPoints = 50;
                        //第一次运行完后，变为非第一行
                        isFirstDataRow = false;
                    }
                    else
                    {
                        sheet.GetRow(i).HeightInPoints = 20;
                    }

                    //设置cell的cellstyle
                    IRow row = sheet.GetRow(i);
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        if (row.GetCell(j) == null)
                        {
                            continue;
                        }
                        //设置单元格样式,必须每一个cell对应一个样式，否则无法设置背景
                        cellStyle = workbook.CreateCellStyle();
                        //垂直居中
                        cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                        //水平对齐
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                        //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                        cellStyle.WrapText = true;

                        //设置cellstyle背景
                        cellStyle.FillBackgroundColor = IndexedColors.Black.Index;
                        cellStyle.FillForegroundColor = IndexedColors.White.Index;

                        //设置边框
                        cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

                        //最后一行设置字体 红色
                        if (i == sheet.LastRowNum)
                        {
                            //创建字体和设置字号
                            IFont font = workbook.CreateFont();
                            font.Color = IndexedColors.Red.Index;
                            cellStyle.SetFont(font);
                        }

                        //填充模式
                        cellStyle.FillPattern = FillPattern.SolidForeground;

                        row.GetCell(j).CellStyle = cellStyle;
                    }
                }

            }
            #endregion

        }

    }
}
