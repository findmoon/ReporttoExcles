using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOIUse;
using NpoiOperatePicture;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

namespace Exicel转换1
{
    class SummaryInfoList
    {
        //使用泛型list存放info信息，信息个数根据job个数确定。
        public List<ExpressSummaryEveryInfo> expressSummaryEveryInfo_list = new List<ExpressSummaryEveryInfo>();

        //dt：Excel表格
        DataTable Resultdt = null;

        #region//应声明为静态字段，查找module时对应需要定位的几个字符串，内容固定
        string findString_machineCount = "Machine count";
        //获取机器种类
        string findString_machineKind = "Machine kind";
        //////获取所有的Module Type,
        string findString_AllModuleType = "Module Type";
        //////获取所有的"Head Type",
        string findString_AllHeadType = "Head Type";
        //获取CPH
        string findString_CPH = "CPH";
        //获取Conveyor,Dual，single
        string findString_Conveyor = "Conveyor";
        //获取 几轨的程式
        string findString_TargetConveyor = "TargetConveyor";
        #endregion
                
        public SummaryInfoList(ExcelHelper excelHelper)
        {
            //ExcelToDataTabl获取sheet到DataTable，第一行是否作为DataTable的列名
            //此处获取dt，是由于其他值的计算需要先获取Result Sheet
            Resultdt = excelHelper.ExcelToDataTable("Result",false);
            //dt = excelHelper.ExcelToDataTable(null, true);            
        }
                             
        // 个位置点
        private int[] Pannel_Point;
        //获取job的位置点，有的report中没有pannel info
        private int[] JOB_Point;
        private int[] Calculation_Point;
        private int[] Machineconf_Point;

        /*
        //  JOB总数,若果有多个程式也记录下来
        int[][] JOBs;
        // pattern 总数
        int[][] calculation_Pattern_POints;
        
        */
        //Machine configration 各个 pattern 的位置
        int[][] machineconf_Pattern_POints;



        #region 获取summaryevery Job info

        #region 获取result下的位置信息

        //获取区间三个位置点    "Pannel information"  "Calculation results"  "Machine configration'  的 位置
        public void getPannelCalculationMachineconf_Point()
        {
            //获取 四个位置点
            string findString_Panel = "Panel information";
            string findString_JOB = "JOB";
            string findString2 = "Calculation results";
            string findString3 = "Machine configration";
            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(Resultdt);
            int[][] Pannel_Points = sectionInfo.GetInfo_Between(findString_Panel);
            int[][] Calculation_Points = sectionInfo.GetInfo_Between(findString2);
            int[][] Machineconf_Points = sectionInfo.GetInfo_Between(findString3);

            Pannel_Point = Pannel_Points[0];
            Calculation_Point = Calculation_Points[0];
            Machineconf_Point = Machineconf_Points[0];
            JOB_Point = sectionInfo.GetInfo_Between(findString_JOB)[0];
        }

        /*
        //获取所有JOB位置
        public void getJOBs()
        {
            //获取的JOB字符
            string findString1 = "JOB";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Pannel_Point, Calculation_Point);
            JOB_Point = sectionInfo.GetInfo_Between(findString_JOB);

        }
        */

        //通用方法，获取指定的Job
        //获取JOB的name\size\Board qty，多个程式时指定获取第几个getTheNum
        public DataTable getJOBDataTable()
        {
            GetCellsDefined getEveryInfoSpecify = new GetCellsDefined();
            //获取Job的DataTable
            ///   1、获取JOB区域,JOB区域 n行 17列
            ///   2、从JOB区域获取获取Board qty位置，获取数量
            /// 第一行均作为DataTable的列，因为sheet生成的表格是个标准的行列对应的表格，便于查找和操作
            /// 
            //Job的DataTable,17列固定，bot有6列名字与top重复，//'''''暂时支取前11列，无bot信息'''''
            DataTable JobDT = getEveryInfoSpecify.GetCellArea(Resultdt, new int[2] { JOB_Point[0] + 1, JOB_Point[1] }, 11, true);
            //
            DataTable jobDT_last=new DataTable();
            jobDT_last.Columns.Add("Name");
            jobDT_last.Columns.Add("Length");
            jobDT_last.Columns.Add("Width"); 
            jobDT_last.Columns.Add("Board qty"); 
            jobDT_last.Columns.Add("Cycle time");
            jobDT_last.Columns.Add("Part qty");

            for (int i = 0; i < JobDT.Rows.Count; i++)
            {
                DataRow jobDR_last = jobDT_last.NewRow();
                jobDR_last["Name"] = JobDT.Rows[i]["Name"];
                jobDR_last["Length"] = JobDT.Rows[i]["Length"];
                jobDR_last["Width"] = JobDT.Rows[i]["Width"];
                jobDR_last["Board qty"] = JobDT.Rows[i]["Board qty"];
                jobDR_last["Cycle time"] = JobDT.Rows[i]["Cycle time"];
                jobDR_last["Part qty"] = JobDT.Rows[i]["Part qty"];
                jobDT_last.Rows.Add(jobDR_last);
            }
            //bottom面的获取
            DataTable JobDT_bottom= getEveryInfoSpecify.GetCellArea(Resultdt, new int[2] { JOB_Point[0] + 1, JOB_Point[1]+ 11 },6 , true);

            //bottom面与top面一致，做到jobDT的数量与后面的表格一致，一一对应
            int hasInsertDRNum = 0;
            for (int i = 0; i < JobDT_bottom.Rows.Count; i++)
            {
                if (Convert.ToInt32(JobDT_bottom.Rows[i]["Part qty"])!=0)
                {

                    DataRow jobDR_last = jobDT_last.NewRow();
                    jobDR_last["Name"] = JobDT.Rows[i]["Name"];
                    jobDR_last["Length"] = JobDT.Rows[i]["Length"];
                    jobDR_last["Width"] = JobDT.Rows[i]["Width"];
                    jobDR_last["Board qty"] = JobDT.Rows[i]["Board qty"];
                    jobDR_last["Cycle time"] = JobDT_bottom.Rows[i]["Cycle time"];
                    jobDR_last["Part qty"] = JobDT_bottom.Rows[i]["Part qty"];

                    //插入到第i+hasInsertDRNum+1行
                    jobDT_last.Rows.InsertAt(jobDR_last, i + hasInsertDRNum + 1);
                    hasInsertDRNum++;
                }
            }
            return jobDT_last;
            
        }

        /*
        //获取Calculation results下的所有 Pattern 位置
        //Pattern1\Pattern2\Pattern3  Pattern格式
        public void getCalculationPatterns()
        {
            //获取的Pattern 字符，不是完全匹配
            string findString1 = "Pattern";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Calculation_Point, Machineconf_Point);

            calculation_Pattern_POints = sectionInfo.GetInfo_Between(findString1, false);

        }
        */

        //获取Machine configration下的所有 Pattern 位置
        //Pattern1\Pattern2\Pattern3  Pattern格式
        public void getMachineconfPatterns()
        {
            //获取的Pattern 字符，不是完全匹配
            string findString_Pattern = "Pattern";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(Resultdt, Machineconf_Point);

            machineconf_Pattern_POints = sectionInfo.GetInfo_Between(findString_Pattern, false);

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
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints);
            }
            else
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints, overThisPoints);
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

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(Resultdt, startString);

            GetCellsDefined getCells = new GetCellsDefined();

            getCells.GetCellArea(Resultdt, sectionInfo.StartPoint, rowX, columnY);


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
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints);
            }
            else
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints, overThisPoints);
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

        #region 判断是否有Bottom面，是否双面pattern
        //判断是否有Customer report字段
        public int[][] getCustomer_reportPoints(int[] underThisPoints,int[] overThisPoints=null)
        {
            string Customer_report = "Customer report";

            int[][] findString_Points;
            //findString_points的位置
            getDataTablePointsInfo findStringPoint;
            //判断有无结束位置
            if (overThisPoints == null)
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints);
            }
            else
            {
                //获取findString_points位置
                findStringPoint = new getDataTablePointsInfo(Resultdt, underThisPoints, overThisPoints);
            }
            findString_Points = findStringPoint.GetInfo_Between(Customer_report);

            if (findString_Points == null)
            {
                //没有找到字符串
                MessageBox.Show("report Excel中没有产生任何报告！请检查确认");
                return null;
            }
            else
            {
                return findString_Points;
            }
            
        }
        #endregion

        #region //获取 生成报告的sheet1各个信息，位置，和最终的summaryeveryifo_List，包含每个job对应的信息。
        public void getEveryInfo()
        {
            //如果打开没有Result的表格，打开其他非报告的表格，则返回DataTable null，重置
            if (Resultdt == null)
            {
                MessageBox.Show(string.Format("未找到{0}表单，请重新选择！", "Result"));
                return;
            }
            getPannelCalculationMachineconf_Point();
            //getJOBs();
            //getCalculationPatterns();
            getMachineconfPatterns();

            
            //Job的DataTable,17列固定，bot有6列名字与top重复，//'''''暂时支取前11列，无bot信息'''''
            DataTable JobDT = getJOBDataTable();

            /*
            //读取calculation Patterns
            //for (int i = 0; i < calculation_Pattern_POints.Rank; i++)
            for (int i = 0; i < 1; i++)
            {
                //MessageBox.Show("calculation_Pattern_POints：" + calculation_Pattern_POints[i][0] + calculation_Pattern_POints[i][1]);
            }
            */



            //读取machineconf Patterns,及其下的 machine count
            //for (int i = 0; i < machineconf_Pattern_POints.Rank; i++),若真循环会不断覆盖值，多个job时需要做到动态添加，或者没获取一个信息则处理一个信心。
            //.Rank获取的数据也不对，.Rank获取的是维数，不是第一维的个数

            //循环获取多个job是对应的信息
            #region //各个信息声明在循环获取的外部，但又不是全局变量
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
            string Pannelsize;
            string Line;
            string Head_Type;
            string jobName;
            string cycleTime;
            string placementNumber;
            string cPH;
            string conveyor;
            string TargetConveyor;
            string remark;
            #endregion

            //拆分开，处理所有Machine configration下面的表格的点
            //位置点变量,以开始point，结束point为一组的方式存放
            List<int[][]> machineconf_Pattern_All_Group_points = new List<int[][]>();
            for (int i = 0; i < machineconf_Pattern_POints.GetLength(0); i++)
            {
                //判断是否双面生产,双面获取双面，单面获取单面
                int[][] Customer_reportPoints = getCustomer_reportPoints(machineconf_Pattern_POints[i], machineconf_Pattern_POints[i + 1]);
                if (Customer_reportPoints == null)
                {//没有任何报告生成
                    continue;
                }
                //获取当前
                if (i< machineconf_Pattern_POints.GetLength(0)-1)
                {
                    int[][] oneSide_point_Group = new int[][] { machineconf_Pattern_POints[i], machineconf_Pattern_POints[i+1] };
                    machineconf_Pattern_All_Group_points.Add(oneSide_point_Group);

                    if (Customer_reportPoints.GetLength(0)==2)
                    {//top面和bottom面的双面
                        int[] anotherSide_startPoint = Customer_reportPoints[1];
                        int[] anotherSide_endPoint=new int[2] { machineconf_Pattern_POints[i + 1][0], Customer_reportPoints[1][1] };
                        int[][] anotherSide_point_Group = new int[][] { anotherSide_startPoint, anotherSide_endPoint };
                        machineconf_Pattern_All_Group_points.Add(anotherSide_point_Group);
                    }
                }
                else
                {
                    //对于最后一个的处理，将结束位置设置为null，getSpecifyString_NextRow会获取开始位置到DataTable结束的位置
                    int[][] oneSide_point_Group = new int[][] { machineconf_Pattern_POints[i],null };
                    machineconf_Pattern_All_Group_points.Add(oneSide_point_Group);

                    if (Customer_reportPoints.GetLength(0) == 2)
                    {//top面和bottom面的双面
                        int[] anotherSide_startPoint = Customer_reportPoints[1];
                        //int[] anotherSide_endPoint = new int[2] { machineconf_Pattern_POints[i + 1][0], Customer_reportPoints[1][1] };
                        int[][] anotherSide_point_Group = new int[][] { anotherSide_startPoint, null};
                        machineconf_Pattern_All_Group_points.Add(anotherSide_point_Group);
                    }

                    #region 111
                    /*
                    //string findString_machineCount = "Machine count";
                    string machineCountString = getSpecifyString_NextRow(findString_machineCount, machineconf_Pattern_POints[i]);
                    machineCount = Convert.ToInt32(machineCountString);
                    // MessageBox.Show("machineCount：" + machineCount);

                    //获取机器种类
                    //string findString_machineKind = "Machine kind";
                    string machineKindString = getSpecifyString_NextRow(findString_machineKind, machineconf_Pattern_POints[i]);
                    machineKind = machineKindString;


                    //获取组合line的信息
                    //////获取所有的Module Type,
                    //string findString_AllModuleType = "Module Type";
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

                    //Try LT-Try统计
                    int count_LT_Tray = 0;
                    int count_LTC_Tray = 0;

                    //拼接Line字符串

                    //重新计算base信息
                    //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
                    //base尽可能少的分配
                    int Count_M = 0;
                    //单独统计M#、M6个数
                    int Count_M3 = 0;
                    int Count_M6 = 0;
                    foreach (var moduleKind in module_Statistics)
                    {
                        if (moduleKind.Key.Substring(0, 2) == "M3")
                        {
                            Count_M += moduleKind.Value;
                            Count_M3 += moduleKind.Value;
                        }
                        if (moduleKind.Key.Substring(0, 2) == "M6")
                        {
                            Count_M += moduleKind.Value * 2;
                            Count_M6 += moduleKind.Value;
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
                    //string findString_AllHeadType = "Head Type";
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
                    //pictureInfoList = new List<PicturesInfo>();
                    //ISheet sheet = excelHelper.ExcelToIsheet("Result");
                    ///int maxColumn = ExcelHelper.sheetColumns(sheet);
                    //仅处理只有一个程式的情况，machineconf_Pattern_POints，到文档结束。
                    //pictureInfoList = NpoiExtend.GetAllPictureInfos(sheet, machineconf_Pattern_POints[i][0], sheet.LastRowNum, machineconf_Pattern_POints[i][1], maxColumn);

                    #endregion

                    //获取CPH
                    //string findString_CPH = "CPH";
                    cPH = getSpecifyString_NextRow(findString_CPH, machineconf_Pattern_POints[i]);

                    //获取Conveyor,Dual，single
                    //string findString_Conveyor = "Conveyor";
                    conveyor = getSpecifyString_NextColumn(findString_Conveyor, machineconf_Pattern_POints[i]);

                    //获取 几轨的程式
                    //string findString_TargetConveyor = "TargetConveyor";
                    TargetConveyor = getSpecifyString_NextColumn(findString_TargetConveyor, machineconf_Pattern_POints[i]);

                    //生成Remake信息
                    //
                    double gonglv = Count_M3 * power_Dictionary["M3III"] + Count_M6 * power_Dictionary["M6III"] +
                        baseCount_2M * power_Dictionary["2MBase"] + baseCount_4M * power_Dictionary["4MBase"] +
                        count_LT_Tray * power_Dictionary["LT-Tray"] + count_LTC_Tray * power_Dictionary["LTC-Tray"];
                    int haoqiliang = baseCount_4M * air_Consumption_Dictionary["4MBase"] + baseCount_2M * air_Consumption_Dictionary["2MBase"];
                    int countu_ip = baseCount_2M + baseCount_4M;
                    double baseLength = baseLength_Dictionary["2MBase"] * baseCount_2M + baseLength_Dictionary["4MBase"] * baseCount_4M;
                    //最后拼接
                    remark = string.Format("功率：{0}\n耗气量：{1}\n长度：{2}\nIP数：{3}",
                        gonglv, haoqiliang, baseLength, countu_ip);
                    */
                    #endregion
                }

            }
            //判断jobDT的数量和machineconf_Pattern_All_Group_points的数量是否一致
            if (JobDT.Rows.Count!= machineconf_Pattern_All_Group_points.Count)
            {
                MessageBox.Show("程序或文件发生了严重的故障！请联系软件作者！！！");
            }

            //根据job的数量 生成summaryinfo
            for (int i = 0; i < JobDT.Rows.Count; i++)
            {
                //每一个的信息
                //job name
                jobName = JobDT.Rows[i]["Name"].ToString();
                //获取"Board qty"
                string BoardQtyString = JobDT.Rows[i]["Board qty"].ToString();
                boardQty = Convert.ToInt32(BoardQtyString);

                //Length Width
                string Length = JobDT.Rows[i]["Length"].ToString();
                string Width = JobDT.Rows[i]["Width"].ToString();
                Pannelsize = Length + "*" + Width;

                //cycletime是line1和 line2的两个时间的平均，但是是在两个报告，两个Excel里面
                //string findString_cycleTime = "Cycle time";
                //string cycleTime_line1 = getSpecifyString_NextRow(findString_cycleTime, JOBs[i], 3, 17);
                //cycleTime = cycleTime_line1;
                string cycleTime_single = JobDT.Rows[i]["Cycle time"].ToString();
                //Part qty
                placementNumber = JobDT.Rows[i]["Part qty"].ToString();

                ///////////////////////////////////-------------------
                //machineconf_Pattern_All_Group_points  成组的开始结束位置

                string machineCountString = getSpecifyString_NextRow(findString_machineCount,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1]);
                machineCount = Convert.ToInt32(machineCountString);
                // MessageBox.Show("machineCount：" + machineCount);


                machineKind = getSpecifyString_NextRow(findString_machineKind,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1]);
                //machineKind = machineKindString;


                //获取组合line的信息
                //获取机器数

                allModuleTypeString = getSpecifyString_NextRow(findString_AllModuleType,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1], machineCount);

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

                //Try LT-Try统计
                int count_LT_Tray = 0;
                int count_LTC_Tray = 0;

                #region//拼接Line字符串

                //重新计算base信息
                //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
                //base尽可能少的分配
                int Count_M = 0;
                //单独统计M#、M6个数
                int Count_M3 = 0;
                int Count_M6 = 0;
                foreach (var moduleKind in module_Statistics)
                {
                    if (moduleKind.Key.Substring(0, 2) == "M3")
                    {
                        Count_M += moduleKind.Value;
                        Count_M3 += moduleKind.Value;
                    }
                    if (moduleKind.Key.Substring(0, 2) == "M6")
                    {
                        Count_M += moduleKind.Value * 2;
                        Count_M6 += moduleKind.Value;
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
                #endregion

                #region//获取组合"Head Type"的信息

                allHeadTypeString = getSpecifyString_NextRow(findString_AllHeadType,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1], machineCount);

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
                #endregion

                #region
                //NPOI导出生成图片，不用判断格式，jpg和png均可生成。
                //获取图片
                //pictureInfoList = new List<PicturesInfo>();
                //ISheet sheet = excelHelper.ExcelToIsheet("Result");
                ///int maxColumn = ExcelHelper.sheetColumns(sheet);
                //仅处理只有一个程式的情况，machineconf_Pattern_POints，到文档结束。
                //pictureInfoList = NpoiExtend.GetAllPictureInfos(sheet, machineconf_Pattern_POints[i][0], sheet.LastRowNum, machineconf_Pattern_POints[i][1], maxColumn);
                //已不从文档中获取图片
                #endregion

                cPH = getSpecifyString_NextRow(findString_CPH,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1]);


                conveyor = getSpecifyString_NextColumn(findString_Conveyor,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1]);


                TargetConveyor = getSpecifyString_NextColumn(findString_TargetConveyor,
                    machineconf_Pattern_All_Group_points[i][0], machineconf_Pattern_All_Group_points[i][1]);

                #region //生成Remake信息
                //
                double gonglv = Count_M3 * ComprehensiveSaticClass.power_Dictionary["M3III"] + 
                    Count_M6 * ComprehensiveSaticClass.power_Dictionary["M6III"] +
                    baseCount_2M * ComprehensiveSaticClass.power_Dictionary["2MBase"] + 
                    baseCount_4M * ComprehensiveSaticClass.power_Dictionary["4MBase"] +
                    count_LT_Tray * ComprehensiveSaticClass.power_Dictionary["LT-Tray"] + 
                    count_LTC_Tray * ComprehensiveSaticClass.power_Dictionary["LTC-Tray"];
                int haoqiliang = baseCount_4M * ComprehensiveSaticClass.air_Consumption_Dictionary["4MBase"] + 
                    baseCount_2M * ComprehensiveSaticClass.air_Consumption_Dictionary["2MBase"];
                int countu_ip = baseCount_2M + baseCount_4M;
                double baseLength = ComprehensiveSaticClass.baseLength_Dictionary["2MBase"] * baseCount_2M + 
                    ComprehensiveSaticClass.baseLength_Dictionary["4MBase"] * baseCount_4M;
                //最后拼接
                remark = string.Format("功率：{0}\n耗气量：{1}\n长度：{2}\nIP数：{3}",
                    gonglv, haoqiliang, baseLength, countu_ip);
                #endregion

                /////由以上信息获取最后的DataTable
                //使用泛型list存放info信息，信息个数根据job个数确定。
                ExpressSummaryEveryInfo expressSummaryEveryInfo = new ExpressSummaryEveryInfo();
                expressSummaryEveryInfo.Jobname = jobName;
                expressSummaryEveryInfo.TargetConveyor = TargetConveyor;
                expressSummaryEveryInfo.Conveyor = conveyor;
                expressSummaryEveryInfo.MachineCount = machineCount;
                //, cPH 两个轨道综合的CPH，由Layoutinfo获取
                expressSummaryEveryInfo.SummaryInfoDT = ComprehensiveSaticClass.getExpressSummayDataTable(boardQty, Line, Head_Type, jobName, Pannelsize, placementNumber, remark);
                expressSummaryEveryInfo.PictureDataByte = ComprehensiveSaticClass.genJobsPicture(allModuleTypeString, module_Statistics);
                expressSummaryEveryInfo.AllModuleType = allModuleTypeString;
                expressSummaryEveryInfo.AllHeadType = allHeadTypeString;

                expressSummaryEveryInfo_list.Add(expressSummaryEveryInfo);
            }
            
        }
        #endregion
        #region 统一处理获取summaryEveryInfo所需要的各个信息，需要多个变量的返回
        // 模组数        
        // 机器种类        
        // 所有的Module Type        
        //所有的HEAD Type
        //字典，存放<模组类型:模组个数>        
        //模组种类，即module_Statistics的key
        ////M3、M6三代、二代\一代的使用
        //string moduleKind3_III = "M3III";
        //string moduleKind6_III = "M6III";
        //string moduleKind3_II = "M3II";
        //string moduleKind6_II = "M6II";
        //string moduleKind3_I = "M3I";
        //string moduleKind6_I = "M6I";
        //字典，存放<HeadType:头个数>       

        //中间图片位置信息        
        //图片信息


        //---生成Excel各个栏的数据
        #endregion
        #endregion

        



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
       
    }


}
