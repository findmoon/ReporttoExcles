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
using NPOI.HSSF.UserModel;
using NPOIUse;
using NpoiOperatePicture;

namespace Exicel转换1
{
    public partial class Form1 : Form
    {
        //ExcelHelper excel处理的类
        ExcelHelper excelHelper;
        //dt：Excel表格
        public DataTable dt=null;
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
        Dictionary<string, int>  module_Statistics ;
        //字典，存放<HeadType:头个数>
        Dictionary<string, int> Head_Statistics;

        //---生成Excel各个栏的数据
        int boardQty;
        string Line;
        string Head_Type;
        string jobName;
        string cycleTime;
        string placementNumber;
        string cPH;
        string conveyor;

        public Form1()
        {
            InitializeComponent();
        }

        #region 打开Excel对话框
        private void ExcelToDTbut_Click(object sender, EventArgs e)
        {
            string file=OpenExcelFile();
            dt = GetDataFromExcle(file);
        }
        #endregion

        #region //打开Excel的方法
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
        public DataTable GetDataFromExcle(string file)
        {
            DataTable dt;
            try
            {
                using (excelHelper = new ExcelHelper(file))
                {
                    dt = excelHelper.ExcelToDataTable("Result", false);
                    //dt = excelHelper.ExcelToDataTable(null, true);
                    return dt;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return null;
            }

        }
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

            Pannel_Point= Pannel_Points[0];
            Calculation_Point= Calculation_Points[0];
            Machineconf_Point= Machineconf_Points[0];
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

            calculation_Pattern_POints = sectionInfo.GetInfo_Between(findString1,false);

        }


        //获取Machine configration下的所有 Pattern 位置
        //Pattern1\Pattern2\Pattern3  Pattern格式
        public void getMachineconfPatterns()
        {
            //获取的Pattern 字符，不是完全匹配
            string findString1 = "Pattern";

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Machineconf_Point);

            machineconf_Pattern_POints = sectionInfo.GetInfo_Between(findString1,false);

        }
#endregion

        #region 一些通用方法

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，从overThisPoints位置之上，根据指定的字符串，获取此字符串下一行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string getSpecifyString_NextRow(string findString, int[] underThisPoints,int[] overThisPoints)
        {
            string[] destString=getSpecifyString_NextRow(findString, underThisPoints, overThisPoints, 1);
            return destString[0];
        }

        /// <summary>
        /// 通用方法 重载，从underThisPoints位置开始，从overThisPoints位置之上，根据指定的字符串，获取此字符串下一行的内容
        /// </summary>
        /// <param name="findString"></param>
        /// <param name="underThisPoints"></param>
        /// <returns></returns>
        public string getSpecifyString_NextRow(string findString, int[] underThisPoints, int RowX,int ColumnY)
        {
            string[] destString = getSpecifyString_NextRow(findString,underThisPoints, new int[] { underThisPoints[0] + RowX, underThisPoints[1] + ColumnY },1);
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
            if (overThisPoints==null)
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

            if (findString_Points==null)
            {
               //没有找到字符串
                return null ;
            }

            //获取findString_points位置下面findN行的内容
            List<string> theDestString = new List<string>();
            for (int i = 0; i < findN; i++)
            {
                int[] lastFindPoints = new int[] { findString_Points[0][0] + (i+1), findString_Points[0][1] };
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
            string[] destString = getSpecifyString_NextRow(findString, underThisPoints,1);
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
            return  getSpecifyString_NextRow(findString, underThisPoints, null, findN);
            
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

        public string getSpecifyString_NextColumn(string findString, int[] underThisPoints, int[] overThisPoints=null)
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
                int[] lastFindPoints = new int[] { findString_Points[0][0] , findString_Points[0][1]+1 };
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
                string findString_cycleTime = "Cycle time";
                string cycleTime_line1 = getSpecifyString_NextRow(findString_cycleTime, JOBs[i], 3, 17);
                cycleTime = cycleTime_line1;

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
                string moduleKind3 = "M3III";
                string moduleKind6 = "M6III";
                module_Statistics.Add(moduleKind3, 0);
                module_Statistics.Add(moduleKind6, 0);

                for (int j = 0; j < allModuleTypeString.Length; j++)
                {
                    if (allModuleTypeString[j] == moduleKind3)
                    {
                        module_Statistics[moduleKind3] += 1;
                    }
                    else if (allModuleTypeString[j] == moduleKind6)
                    {
                        module_Statistics[moduleKind6] += 1;
                    }

                }

                //拼接Line字符串
                //计算base信息
                int baseCount_4M = ((module_Statistics[moduleKind3] + module_Statistics[moduleKind6] * 2) / 4);
                int baseCount_2M;
                if (((module_Statistics[moduleKind3] + module_Statistics[moduleKind6] * 2) % 4) != 0)
                {
                    baseCount_2M = 1;
                }
                else
                {
                    baseCount_2M = 0;
                }

                //"/r/n" 回车换行符
                Line = "";
                Line = machineKind + "-(" + moduleKind3 + "*" + module_Statistics[moduleKind3].ToString() + "+" + moduleKind6 + "*" + module_Statistics[moduleKind6] + ")" + "/r/n" + "4M BASE III*" + baseCount_4M.ToString() + "+2M BASE III*" + baseCount_2M.ToString();
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
                ////NPOI导出生成图片，不用判断格式，jpg和png均可生成。
                ////获取图片
                //List<PicturesInfo> pictureInfoList = new List<PicturesInfo>();
                //ISheet sheet = excelHelper.ExcelToIsheet("Result");
                //int maxColumn = excelHelper.sheetColumns(sheet);
                ////仅处理只有一个程式的情况，machineconf_Pattern_POints，到文档结束。
                //pictureInfoList=NpoiExtend.GetAllPictureInfos(sheet, machineconf_Pattern_POints[i][0], sheet.LastRowNum,machineconf_Pattern_POints[i][1], maxColumn);

                ////测试>>>生成图片表格
                //string file1 = @"D:\Test\120.xlsx";

                ////ExcelHelper类内部创建
                ////if (!File.Exists(file1))
                ////{
                ////    /*
                ////    //file1正在被另一进程使用，，
                ////    File.Create(file1);
                ////    */
                ////    using (File.Create(file1))
                ////    {

                ////    }
                ////}

                //ExcelHelper excelHelper1 = new ExcelHelper(file1);

                //ISheet sheet1 = excelHelper1.ExcelToIsheet();

                //int startRow;
                //int startCol;
                ////pictureInfoList需要排序
                //foreach (PicturesInfo pictureInfo in pictureInfoList)
                //{
                //    /*
                //    using (FileStream fsPicture=new FileStream(@"D:\Test\"+pictureInfoList.IndexOf(item)+".png",FileMode.OpenOrCreate,FileAccess.ReadWrite))
                //    {
                //        fsPicture.Write(pictureInfo.PictureData,0, pictureInfo.PictureData.Length);
                //    }
                //    */



                //        if (pictureInfoList.IndexOf(pictureInfo) == 0)
                //        {
                //            startRow = 2;
                //            startCol = 2;
                //        }
                //        else
                //        {
                //            startRow = 2;
                //            //startRow = 2 + pictureInfoList[pictureInfoList.IndexOf(pictureInfo) - 1].MaxRow - pictureInfoList[pictureInfoList.IndexOf(pictureInfo) - 1].MinRow;
                //            startCol = 2 + pictureInfoList[pictureInfoList.IndexOf(pictureInfo) - 1].MaxCol - pictureInfoList[pictureInfoList.IndexOf(pictureInfo) - 1].MinCol;
                //        }



                //        excelHelper1.pictureDataToSheet(sheet1, pictureInfo.PictureData, startRow, startCol, startRow + pictureInfo.MaxRow - pictureInfo.MinRow, startCol + pictureInfo.MaxCol - pictureInfo.MinCol);



                //    //MessageBox.Show("OK");
                //}

                //using (FileStream file1fs =new FileStream(file1,FileMode.OpenOrCreate,FileAccess.ReadWrite) ) 
                //{
                //    excelHelper1.WorkBook.Write(file1fs);
                //}

                #endregion


                //获取CPH
                string findString_CPH = "CPH";
                cPH = getSpecifyString_NextRow(findString_CPH, machineconf_Pattern_POints[i]);

                //获取Conveyor
                string findString_Conveyor = "Conveyor";
                conveyor = getSpecifyString_NextColumn(findString_Conveyor, machineconf_Pattern_POints[i]);
            }



        }
        #endregion


  

        //生成DataTable
        //生成评估sheet1
        private void GenExlbut_Click(object sender, EventArgs e)
        {
            //保存的路径
            //string fileName = @"D:\Test\110.xlsx";
            string fileName = null;
            saveExcelFileDialog.Filter = "Excel 2007 工作表(*.xlsx)|*.xlsx|Excel 97-2003 工作表(*.xls)|*.xls";
            saveExcelFileDialog.RestoreDirectory = true;
            if (saveExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveExcelFileDialog.FileName;
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

            //创建ExcelHelper，用以存放sheet1和sheet2
            ExcelHelper excelHelper1 = new ExcelHelper(fileName);

            #region 第一个sheet的获取和生成
            //第一个sheet的标题，以字符串形式插入
            string oneTitle = "FUJI Line Throughput Report 三代";

            DataTable genDT = getFirstSheetDataTable();

            string sheet1 = "sheet1";

            //插入sheet对应表格的datatable
            DataTableToResultSheet1(excelHelper1, genDT, sheet1, true, new int[2] { 3, 1 });
            //插入标题
            StringInsertExcel(excelHelper1, oneTitle, sheet1, new int[2] { 1, 1 });
            #endregion


            #region 第二个sheet的获取和生成
            //第二个sheet的标题，以字符串形式插入
            string secondTitle = jobName;

            //
            //获取生成的第二个sheet对象在Excel对应的sheetname
            string sheetName_CustomerReport = "CustomerReport_1T";

            DataTable genDT2 = getSecondSheetDataTable(sheetName_CustomerReport);
            

            string sheet2 = "sheet2";

            //插入sheet对应表格的datatable
            DataTableToResultSheet2(excelHelper1, genDT2, sheet2, true, new int[2] { 3, 2 });
            //插入标题
            StringInsertExcel(excelHelper1, secondTitle, sheet2, new int[2] { 2, 2 });
            #endregion


            //获取对象excelHelper1的文件流，即使用filename创建的文件流，判断是否有文件，没有文件的fa为null
            FileStream fs = excelHelper1.FS;
            if (fs == null)
            {
                fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            }
            excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流

            MessageBox.Show("生成Excel成功！");
        }

        #region //根据各种信息生成第一个sheet信息的DataTable
        public DataTable  getFirstSheetDataTable()
        {
            getEveryInfo();

           

            //生成评估的表格
            DataTable genDT = new DataTable();
            string[] genColumnString = { "Item", "Board", "Line" , "Head Type","Job name /r/n Panel Size" ,"Cycle time /n [((L1+L2)/2) Sec]", "Number of  placements", "CPH", "Output Board/hour" , "Output Board/Day(22H)", "CPH Rate", "Remark" };

            //评估表格的栏
            for (int i = 0; i < genColumnString.Length; i++)
            {
                genDT.Columns.Add(genColumnString[i]);
            }

            //像生成的Excel添加行
            DataRow genRow = genDT.NewRow();
            genRow["Board"] = boardQty;
            genRow["Line"] = Line;
            genRow["Head Type"] = Head_Type;
            genRow["Job name /r/n Panel Size"] = jobName;
            genRow["Cycle time /n [((L1+L2)/2) Sec]"] = cycleTime;
            genRow["Number of  placements"] = placementNumber;
            genRow["CPH"] = cPH;
            //float.Parse(cycleTime)  float转换
            double double1 = Convert.ToDouble(boardQty) * 3600 / Convert.ToDouble(cycleTime);
            genRow["Output Board/hour"] = string.Format("{0:0}", double1);
            genRow["Output Board/Day(22H)"] = string.Format("{0:0}", double1 * 22);
            //42000  H24头理论的CPH，10500  H04SF理论的CPH
            //需要建一个head type:CPH的字典，进行计算；
            double CPHRate = Convert.ToDouble(cPH) / (42000 * 7 + 10500 * 1);
            genRow["CPH Rate"] = string.Format("{0:0.00%}",CPHRate);
            genRow["Remark"] = "已删减大零件";

            //添加行
            genDT.Rows.Add(genRow);

            return genDT;

           

            


        }
        #endregion

        #region //根据各种信息生成第二个sheet信息的DataTable
        public DataTable getSecondSheetDataTable(string sheetName_CustomerReport)
        {
            #region //获取生成的第二个sheet对象的sheetname
            
            getsecondsheet secondsheet = new getsecondsheet(excelHelper, sheetName_CustomerReport, jobName, conveyor);

            //获取二轨的Excel生成对象



            //最终要组合的行
            int secondSheetrowCount = secondsheet.SmallSecondSheetDT.Rows.Count > secondsheet.SecondSheetDT.Rows.Count ? secondsheet.SmallSecondSheetDT.Rows.Count : secondsheet.SecondSheetDT.Rows.Count;
            DataTable secondSheetDT = new DataTable();

            //初始化secondSheetDT的column
            //最终初始化的栏  i = 0; i < 15; i++)

            //设置secondSheetDT的DataTable的第一行

            secondSheetDT.Columns.Add("Prouction Mode");
            secondSheetDT.Columns.Add("Module No.");
            secondSheetDT.Columns.Add("Module");
            secondSheetDT.Columns.Add("Head Type");
            secondSheetDT.Columns.Add("Lane1 Cycle Time");
            secondSheetDT.Columns.Add("Lane2 Cycle Time");
            secondSheetDT.Columns.Add("Cycle Time");
            secondSheetDT.Columns.Add("Qty");
            secondSheetDT.Columns.Add("Avg");
            secondSheetDT.Columns.Add("CPH");

            secondSheetDT.Columns.Add("Feeder");
            secondSheetDT.Columns.Add("Feeder Qty");
            secondSheetDT.Columns.Add("Head Type ");
            secondSheetDT.Columns.Add("Nozzle");
            secondSheetDT.Columns.Add("Nozzle Qty");

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
                //lin2 cycyle time
                //secondSheetDT.Rows[i][5] = secondsheet2.SecondSheetDT.Rows[i][4];
                //平均cycle time
                //secondSheetDT.Rows[i][6] = (secondsheet.SecondSheetDT.Rows[i][4]+ secondsheet2.SecondSheetDT.Rows[i][4])/2;
                //Qty
                secondSheetDT.Rows[i][7] = secondsheet.SecondSheetDT.Rows[i][5];
                //Avg  问题 Cycle time/Qty
                //secondSheetDT.Rows[i][8] = System.Convert.ToDouble(secondSheetDT.Rows[i][6])/ System.Convert.ToDouble(secondSheetDT.Rows[i][7]);
                //CPH  问题 Qty*3600/CyCle time
                //secondSheetDT.Rows[i][9] = System.Convert.ToInt32(secondSheetDT.Rows[i][7])*3600/System.Convert.ToInt32(secondSheetDT.Rows[i][6]);

                //Feeder
                secondSheetDT.Rows[i][10] = secondsheet.SmallSecondSheetDT.Rows[i][0];
                secondSheetDT.Rows[i][11] = secondsheet.SmallSecondSheetDT.Rows[i][1];
                secondSheetDT.Rows[i][12] = secondsheet.SmallSecondSheetDT.Rows[i][2];
                secondSheetDT.Rows[i][13] = secondsheet.SmallSecondSheetDT.Rows[i][3];
                secondSheetDT.Rows[i][14] = secondsheet.SmallSecondSheetDT.Rows[i][4];
            }

            //计算最后的平均值
            secondSheetrowCount++;
            DataRow dr2 = secondSheetDT.NewRow();
            secondSheetDT.Rows.Add(dr2);
            secondSheetDT.Rows[secondSheetrowCount - 1][0] = "Total";
            secondSheetDT.Rows[secondSheetrowCount - 1][4] = "";
            secondSheetDT.Rows[secondSheetrowCount - 1][5] = "";
            secondSheetDT.Rows[secondSheetrowCount - 1][6] = "";
            secondSheetDT.Rows[secondSheetrowCount - 1][7] = "";
            secondSheetDT.Rows[secondSheetrowCount - 1][8] = "";
            secondSheetDT.Rows[secondSheetrowCount - 1][9] = "";

            //验证获取到的DataTable
            return secondSheetDT;
            #endregion
        }
        #endregion



        //将string插入Excel指定位置
        public void StringInsertExcel(ExcelHelper excelHelper1, string insertString,string sheetName,int[] insertPoint)
        {
            //使用指定的fileName 创建Iworkbook


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

                XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                //垂直居中
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //水平对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                //应用到单元格样式
                row.GetCell(insertPoint[1]).CellStyle = cellStyle;

                #region 写入文件流应该在方法外，所有插入的数据和样式设置好后写入，并关闭流
                //fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                //workbook.Write(fs); //写入到excel文件，并关闭流
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
        public void DataTableToResultSheet1(ExcelHelper excelHelper1,DataTable dt, string sheetName,bool isColumnWritten, int[] insertPoint)
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
                    count ++;
                }


                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1]+j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                    ++count;
                }


                #region sheet1的样式
                //设置列宽
                int commonWidth = 10;
                sheet.SetColumnWidth(0, 1 * 256);
                sheet.SetColumnWidth(1, commonWidth * 256);
                sheet.SetColumnWidth(2, commonWidth * 256);
                sheet.SetColumnWidth(3, System.Convert.ToInt32(System.Convert.ToDouble(commonWidth * 256)*1.5));
                sheet.SetColumnWidth(4, commonWidth * 2 * 256);
                sheet.SetColumnWidth(5, commonWidth * 256);
                sheet.SetColumnWidth(6, commonWidth * 256);
                sheet.SetColumnWidth(7, commonWidth * 256);
                sheet.SetColumnWidth(8, commonWidth * 256);
                sheet.SetColumnWidth(9, commonWidth * 256);
                sheet.SetColumnWidth(10, commonWidth * 256);
                sheet.SetColumnWidth(11, commonWidth * 256);
                sheet.SetColumnWidth(12, commonWidth * 256);


                //合并单元格
                sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 13));//2.0使用 2.0以下为Region
                sheet.AddMergedRegion(new CellRangeAddress(2, 2, 1, 13));//2.0使用 2.0以下为Region

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
                row2.HeightInPoints = 60;
                row3.HeightInPoints = 30;
                row4.HeightInPoints = 50;


                //设置单元格样式

                XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                //垂直居中
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //水平对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                //设置整列单元格样式
                //获取列数量
                int column_count = ExcelHelper.sheetColumns(sheet);
                for (int i = 0; i < column_count; i++)
                {
                    sheet.SetDefaultColumnStyle(i, cellStyle);
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


                setSecondSheetStyle(sheet, dt.Columns.Count);


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



        public void setSecondSheetStyle(ISheet sheet,int mergeLastCount)
        {
            //应使用静态方法，可以获取不规则行列的sheet的最大与最小列
            int sheetColumns = ExcelHelper.sheetColumns(sheet);
            #region sheet2的样式
            //设置列宽
            int commonWidth = 15;
            sheet.SetColumnWidth(0, 1 * 256);
            for (int i = 0; i < sheetColumns; i++)
            {
                sheet.SetColumnWidth(12, commonWidth * 256);
            }
            


            ////合并单元格
            //sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 13));//2.0使用 2.0以下为Region
            //sheet.AddMergedRegion(new CellRangeAddress(2, 2, 1, 13));//2.0使用 2.0以下为Region

            //设置行高
            bool isFirstDataRow = true;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                if (sheet.GetRow(i) == null)
                {
                        sheet.CreateRow(i).HeightInPoints = 20;
                }
                else
                {
                    if (isFirstDataRow)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(i - 1, i - 1, 2, 2+mergeLastCount));

                        sheet.GetRow(i).HeightInPoints = 60;
                    }
                    else
                    {                        
                        sheet.GetRow(i).HeightInPoints = 20;
                    }
                    
                    isFirstDataRow = false;
                }

                
            }

  

            //设置单元格样式

            XSSFCellStyle cellStyle = (XSSFCellStyle)sheet.Workbook.CreateCellStyle();
            //垂直居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            //水平对齐
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

            //设置整列单元格样式
            //获取列数量

            for (int i = 0; i < sheetColumns; i++)
            {
                sheet.SetDefaultColumnStyle(i, cellStyle);
            }
            #endregion

        }








    }
}
