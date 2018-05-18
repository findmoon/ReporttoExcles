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


namespace Exicel转换1
{
    class layoutInfoList
    {
        //存放获取的CustomerRepor的sheet列表
        List<ISheet> sheet_CustomerReportList = new List<ISheet>();

        //layoutExpressEveryInfo
        public List<layoutExpressEveryInfo> layoutExpressEveryInfo_list = new List<layoutExpressEveryInfo>();

        

        

        //layoutInfoList构造函数,初始化sheet_CustomerReportList
        public layoutInfoList(ExcelHelper excelHelper)
        {
            //获取excelHelper下的各个_CustomerReport

            int sheetNum = excelHelper.WorkBook.NumberOfSheets;
            for (int i = 0; i < sheetNum; i++)
            {
                if (excelHelper.WorkBook.GetSheetName(i).Contains("CustomerReport"))
                {
                    sheet_CustomerReportList.Add(excelHelper.WorkBook.GetSheetAt(i));
                }
                
            }
        }
        
        
        //生成第二个表格所需要的各个信息
        //返回0，生成失败，返回1生成成功
        public bool getEveryInfo()
        {
            DataTable CustomerReport_DataTable = null;

            //固定查找字符串放在外部
            string detailString = "Detail";
            string feederNumberString = "Feeder required number";
            string NozzleNumberString = "Nozzle required number";
            string AutoToolString = "AutoTool required number";

            DataTable feederNozzleDT;
            DataTable layoutExpressDT;

            for (int i = 0; i < sheet_CustomerReportList.Count; i++)
            {
                //获取当前的CustomerReport sheet到DataTable
                CustomerReport_DataTable = ExcelHelper.IsheetToDataTable(sheet_CustomerReportList[i]);
                if (CustomerReport_DataTable == null)
                {
                    MessageBox.Show(string.Format("sheet：{0},未获取成功!，请重新打开刚才的文件" + sheet_CustomerReportList[i]));
                    return false;
                }

                int[] zorePoint = new int[] { 0, 0 };
                //获取jobName
                string jobName = getSpecifyString_NextColumn(CustomerReport_DataTable, "Name", zorePoint);

                //获取是top还是bot
                string t_or_b = getSpecifyString_NextColumn(CustomerReport_DataTable, "T/B", zorePoint);

                //获取Machine kind，判断是一段还是多段，NXTIII，或NXTIII | NXTII，
                //获取Conveyor，
                string machineKind = getSpecifyString_NextColumn(CustomerReport_DataTable, "Machine kind", zorePoint);
                string Conveyor = getSpecifyString_NextColumn(CustomerReport_DataTable, "Conveyor", zorePoint);
                //处理dual,-，dual,dual，single，duoble等情况
                if (Conveyor.Contains("Dual"))
                {
                    Conveyor = "Dual";
                }
                else if (Conveyor.Contains("Single"))
                {
                    Conveyor = "Single";
                }
                else if (Conveyor.Contains("Double"))
                {
                    Conveyor = "Double";
                }


                #region 获取Detail，Feeder required number，Nozzle required number位置
                //获取Detail，Feeder required number，Nozzle required number位置
                int[] detailPoint = getPoints(CustomerReport_DataTable, detailString);
                int[] feederNumberPoint = getPoints(CustomerReport_DataTable, feederNumberString);
                int[] NozzleNumberPoint = getPoints(CustomerReport_DataTable, NozzleNumberString);
                int[] AutoToolPoint = getPoints(CustomerReport_DataTable, AutoToolString);
                #endregion

                #region 获取上面两点间某一位置的数据，因为sheetName_CustomerReport 只会有单个程式
                //获取Module
                string moduleString = "Module";
                int[] modulePoint = getPoints(CustomerReport_DataTable, detailPoint, feederNumberPoint, moduleString);

                string feederString = "Feeder";
                int[] feederPoint = getPoints(CustomerReport_DataTable, feederNumberPoint, NozzleNumberPoint, feederString);

                string nozzleNameString = "NozzleName";
                int[] nozzleNamePoint = getPoints(CustomerReport_DataTable, NozzleNumberPoint, AutoToolPoint, nozzleNameString);
                #endregion


                #region 获取确定点的区域
                //第一行均作为DataTable的列，因为sheet生成的表格是个标准的行列对应的表格，便于查找和操作
                //获取module区域，9行，12列的区域, 行=模组数 machineCount+1 ，12列固定
                //int[] moduleArea = new int[] { machineCount + 1, 12 };
                GetCellsDefined getSomeCells = new GetCellsDefined();

                //获取"Module"下的datat
                DataTable moduleDT = getSomeCells.GetCellArea(CustomerReport_DataTable, modulePoint, 12, true);


                //获取Feeder区域
                //feeder区域结束位置不知，由空值决定结束位置，2列
                DataTable feederDT = getSomeCells.GetCellArea(CustomerReport_DataTable, feederPoint, 2, true);
                //获取这个区域的DataTable


                //获取NozzleName区域
                //feeder区域结束位置依旧不知，由空值决定结束位置，2列
                DataTable nozzleNameDT = getSomeCells.GetCellArea(CustomerReport_DataTable, nozzleNamePoint, 2, true);
                //获取这个区域的DataTable
                #endregion

                #region 获取生成layoutinfo layexpressDT和feederNozzleDT的数据，totalfeederNozzleDT无
                //获取，最终生成表的数据
                //Prouction Mode \ Module No. \ Module\Head Type\Lane1 Cycle Time\
                //Lane2 Cycle Time  需要另外一个文件生成2轨的cycle time
                // Cycle Time  两个轨道时间平均值
                //Qty  生产总个数
                //Avg  每个所用的时间
                //CPH  Qty*3600/Cycle Time
                //Feeder 
                //Qty Feeder的数量
                //Head  Type
                //Nozzle
                //QTy  Nozzle的数量
                layoutExpressDT = new DataTable();
                //小表格
                feederNozzleDT = new DataTable();

                //初始化各个栏
                layoutExpressDT.Columns.Add("Prouction Mode");
                layoutExpressDT.Columns.Add("Module No.");
                layoutExpressDT.Columns.Add("Module");
                layoutExpressDT.Columns.Add("Head Type");
                layoutExpressDT.Columns.Add("Cycle Time");
                layoutExpressDT.Columns.Add("Qty");

                feederNozzleDT.Columns.Add("Feeder");
                feederNozzleDT.Columns.Add("Qty Feeder");
                feederNozzleDT.Columns.Add("Head  Type");
                feederNozzleDT.Columns.Add("Nozzle");
                feederNozzleDT.Columns.Add("QTy  Nozzle");


                int feederNozzleDTRowCount = nozzleNameDT.Rows.Count > feederDT.Rows.Count ? nozzleNameDT.Rows.Count : feederDT.Rows.Count;
                int last_layoutExpressDTRowCount = feederNozzleDTRowCount > moduleDT.Rows.Count ? feederNozzleDTRowCount : moduleDT.Rows.Count;
                //循环获取各个单元格的值
                for (int j = 0; j < last_layoutExpressDTRowCount; j++)
                {
                    DataRow dr = layoutExpressDT.NewRow();

                    if (j<moduleDT.Rows.Count)
                    {
                        dr["Prouction Mode"] = Conveyor;
                    }
                    /*//空行的特殊处理
                    if (j >= moduleDT.Rows.Count)
                    {
                        DataRow dr1 = moduleDT.NewRow();
                        moduleDT.Rows.Add(dr1);
                        dr["Prouction Mode"] = null;
                    }
                    */


                    dr["Module No."] = moduleDT.Rows[j]["Module"];
                    dr["Module"] = moduleDT.Rows[j]["Module Type"];
                    dr["Head Type"] = moduleDT.Rows[j]["Head Type"];
                    dr["Cycle Time"] = moduleDT.Rows[j]["CycleTime"];
                    dr["Qty"] = moduleDT.Rows[j]["Qty"];

                    //将row添加到DataTable
                    layoutExpressDT.Rows.Add(dr);
                }


                for (int j = 0; j < last_layoutExpressDTRowCount; j++)
                {
                    if (nozzleNameDT.Rows.Count <= j)
                    {
                        DataRow dr1 = nozzleNameDT.NewRow();
                        nozzleNameDT.Rows.Add(dr1);
                    }
                    if (feederDT.Rows.Count <= j)
                    {
                        DataRow dr1 = feederDT.NewRow();
                        feederDT.Rows.Add(dr1);
                    }

                    DataRow dr = feederNozzleDT.NewRow();
                    dr["Feeder"] = feederDT.Rows[j]["Feeder"];
                    dr["Qty Feeder"] = feederDT.Rows[j]["Qty"];
                    dr["Nozzle"] = nozzleNameDT.Rows[j]["NozzleName"];
                    dr["QTy  Nozzle"] = nozzleNameDT.Rows[j]["Qty"];

                    //head type的单独添加
                    if (j < moduleDT.Rows.Count)
                    {
                        dr["Head  Type"] = moduleDT.Rows[j]["Head Type"];
                    }

                    feederNozzleDT.Rows.Add(dr);
                }
                #endregion

                layoutExpressEveryInfo layoutExpressEveryInfo = new layoutExpressEveryInfo();
                layoutExpressEveryInfo.jobName = jobName;
                layoutExpressEveryInfo.t_or_b = t_or_b;
                layoutExpressEveryInfo.layoutExpressDT = layoutExpressDT;
                layoutExpressEveryInfo.feederNozzleDT = feederNozzleDT;

                layoutExpressEveryInfo_list.Add(layoutExpressEveryInfo);
            }

            return true;

            #region 获取sheetname_Result的各种信息
            //this.sheetName_Result = sheetName_Result;
            //DataTable Result_DataTable = null;
            ////ExcelToDataTabl获取sheet到DataTable，第一行是否作为DataTable的列名
            //Result_DataTable = excelHelper.ExcelToDataTable(sheetName_Result, false);
            //if (Result_DataTable == null)
            //{
            //    MessageBox.Show("未获取到sheet：" + CustomerReport_DataTable);
            //    return 0;
            //}

            //int[] targetConveyor_Point = getPoints(Result_DataTable, "TargetConveyor");

            #endregion
        }

        #region//getsecondsheet构造函数
        /*
         * private string JOBName;
        private string Conveyor;
        private ExcelHelper excelHelper = null;
        private string sheetName_CustomerReport;

        private int machineCount;
        private string sheetName_Result;
        private string Job;
        private string targetConveyor;
        public getLayoutDT(ExcelHelper excelHelper, string sheetName_CustomerReport, string JOBName, string Conveyor, int machineCount)
        {
            #region 获取sheetName_CustomerReport的各种信息
            this.JOBName = JOBName;
            this.Conveyor = Conveyor;
            this.excelHelper = excelHelper;
            this.sheetName_CustomerReport = sheetName_CustomerReport;
            this.machineCount = machineCount;
        }
        
        //生成第二个表格所需要的各个信息
        //返回0，生成失败，返回1生成成功
        public bool getsecondsheetInfo()
        {

            DataTable CustomerReport_DataTable = null;
            //ExcelToDataTabl获取sheet到DataTable，第一行是否作为DataTable的列名
            CustomerReport_DataTable = excelHelper.ExcelToDataTable(sheetName_CustomerReport, false);
            if (CustomerReport_DataTable == null)
            {
                MessageBox.Show("未获取到sheet：" + CustomerReport_DataTable);
                return false;
            }

            //判断jobname，前后选择的两个报告是否一致
            //int[] find_confirmJobNamePoint = getPoints(CustomerReport_DataTable, "Name");
            string confirmJobName = getSpecifyString_NextColumn(CustomerReport_DataTable, "Name", new int[] { 0, 0 }, new int[] { 10, 3 });
            if (confirmJobName != this.JOBName)
            {
                MessageBox.Show("所选两个报告的Job名字不一样，无法生成，请重新选择同一程式两个轨道对应的生产报告！");
                return false;
            }

            #region 获取Detail，Feeder required number，Nozzle required number位置
            //获取Detail，Feeder required number，Nozzle required number位置
            string detailString = "Detail";
            string feederNumberString = "Feeder required number";
            string NozzleNumberString = "Nozzle required number";
            string AutoToolString = "AutoTool required number";

            int[] detailPoint = getPoints(CustomerReport_DataTable, detailString);
            int[] feederNumberPoint = getPoints(CustomerReport_DataTable, feederNumberString);
            int[] NozzleNumberPoint = getPoints(CustomerReport_DataTable, NozzleNumberString);
            int[] AutoToolPoint = getPoints(CustomerReport_DataTable, AutoToolString);
            #endregion

            #region 获取上面两点间某一位置的数据，因为sheetName_CustomerReport 只会有单个程式
            //获取Module
            string moduleString = "Module";
            int[] modulePoint = getPoints(CustomerReport_DataTable, detailPoint, feederNumberPoint, moduleString);

            string feederString = "Feeder";
            int[] feederPoint = getPoints(CustomerReport_DataTable, feederNumberPoint, NozzleNumberPoint, feederString);

            string nozzleNameString = "NozzleName";
            int[] nozzleNamePoint = getPoints(CustomerReport_DataTable, NozzleNumberPoint, AutoToolPoint, nozzleNameString);
            #endregion

            #region 获取确定点的区域
            //第一行均作为DataTable的列，因为sheet生成的表格是个标准的行列对应的表格，便于查找和操作
            //获取module区域，9行，12列的区域, 行=模组数 machineCount+1 ，12列固定
            int[] moduleArea = new int[] { machineCount + 1, 12 };
            GetCellsDefined getModuleCells = new GetCellsDefined();
            getModuleCells.GetCellArea(CustomerReport_DataTable, modulePoint, moduleArea[0], moduleArea[1], true);
            //获取这个区域的DataTable
            DataTable moduleDT = getModuleCells.DT;

            //获取Feeder区域
            //feeder区域结束位置不知，暂值给定Nozzle required number--- NozzleNumberPoint的行-1为行个数，2列
            int feederRowCount = NozzleNumberPoint[0] - feederPoint[0] - 1;
            int[] feederArea = new int[] { feederRowCount, 2 };
            GetCellsDefined getFeederCells = new GetCellsDefined();

            getFeederCells.GetCellArea(CustomerReport_DataTable, feederPoint, feederArea[0], feederArea[1], true);
            //获取这个区域的DataTable
            DataTable feederDT = getFeederCells.DT;


            //获取NozzleName区域
            //feeder区域结束位置依旧不知，暂值给定AutoTool required number--- AutoToolPoint的行-1为行个数，2列
            int nozzleNameRowCount = AutoToolPoint[0] - nozzleNamePoint[0] - 1;
            int[] nozzleNameArea = new int[] { nozzleNameRowCount, 2 };
            GetCellsDefined getnozzleNameCells = new GetCellsDefined();
            getnozzleNameCells.GetCellArea(CustomerReport_DataTable, nozzleNamePoint, nozzleNameArea[0], nozzleNameArea[1], true);
            //获取这个区域的DataTable
            DataTable nozzleNameDT = getnozzleNameCells.DT;
            #endregion

            #region 获取申城sheet区域的数据
            //获取，最终生成表的数据
            //Prouction Mode \ Module No. \ Module\Head Type\Lane1 Cycle Time\
            //Lane2 Cycle Time  需要另外一个文件生成2轨的cycle time
            // Cycle Time  两个轨道时间平均值
            //Qty  生产总个数
            //Avg  每个所用的时间
            //CPH  Qty*3600/Cycle Time
            //Feeder 
            //Qty Feeder的数量
            //Head  Type
            //Nozzle
            //QTy  Nozzle的数量
            secondSheetDT = new DataTable();
            //小表格
            smallsecondSheetDT = new DataTable();

            //初始化各个栏
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


            int smallsecondSheetDTRowCount = nozzleNameDT.Rows.Count > feederDT.Rows.Count ? nozzleNameDT.Rows.Count : feederDT.Rows.Count;
            int secondSheetDTRowCount = smallsecondSheetDTRowCount > moduleDT.Rows.Count ? smallsecondSheetDTRowCount : moduleDT.Rows.Count;
            //循环获取各个单元格的值
            for (int i = 0; i < secondSheetDTRowCount; i++)
            {
                DataRow dr = secondSheetDT.NewRow();

                dr["Prouction Mode"] = Conveyor;

                //空行的特殊处理
                if (i >= moduleDT.Rows.Count)
                {
                    DataRow dr1 = moduleDT.NewRow();
                    moduleDT.Rows.Add(dr1);
                    dr["Prouction Mode"] = null;
                }



                dr["Module No."] = moduleDT.Rows[i]["Module"];
                dr["Module"] = moduleDT.Rows[i]["Module Type"];
                dr["Head Type"] = moduleDT.Rows[i]["Head Type"];
                dr["Cycle Time"] = moduleDT.Rows[i]["CycleTime"];
                dr["Qty"] = moduleDT.Rows[i]["Qty"];

                //将row添加到DataTable
                secondSheetDT.Rows.Add(dr);
            }


            for (int i = 0; i < secondSheetDTRowCount; i++)
            {
                if (nozzleNameDT.Rows.Count <= i)
                {
                    DataRow dr1 = nozzleNameDT.NewRow();
                    nozzleNameDT.Rows.Add(dr1);
                }
                if (feederDT.Rows.Count <= i)
                {
                    DataRow dr1 = feederDT.NewRow();
                    feederDT.Rows.Add(dr1);
                }

                DataRow dr = smallsecondSheetDT.NewRow();
                dr["Feeder"] = feederDT.Rows[i]["Feeder"];
                dr["Qty Feeder"] = feederDT.Rows[i]["Qty"];
                dr["Nozzle"] = nozzleNameDT.Rows[i]["NozzleName"];
                dr["QTy  Nozzle"] = nozzleNameDT.Rows[i]["Qty"];

                //head type的单独添加
                if (i < moduleDT.Rows.Count)
                {
                    dr["Head  Type"] = moduleDT.Rows[i]["Head Type"];
                }


                smallsecondSheetDT.Rows.Add(dr);
            }
            #endregion
            #endregion

            return true;

            #region 获取sheetname_Result的各种信息
            //this.sheetName_Result = sheetName_Result;
            //DataTable Result_DataTable = null;
            ////ExcelToDataTabl获取sheet到DataTable，第一行是否作为DataTable的列名
            //Result_DataTable = excelHelper.ExcelToDataTable(sheetName_Result, false);
            //if (Result_DataTable == null)
            //{
            //    MessageBox.Show("未获取到sheet：" + CustomerReport_DataTable);
            //    return 0;
            //}

            //int[] targetConveyor_Point = getPoints(Result_DataTable, "TargetConveyor");

            #endregion
        }
        */
        #endregion

        #region 普通通用方法，获取DataTable中某一string的位置
        //获取DataTable中某一string的位置
        public int[] getPoints(DataTable dt, string findString)
        {

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt);
            int[][] Pannel_Points = sectionInfo.GetInfo_Between(findString);

            return Pannel_Points[0];
        }

        //获取DataTable中两点间某一string的位置
        public int[] getPoints(DataTable dt, int[] Point1, int[] Point2, string findString1)
        {
            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt, Point1, Point2);
            return sectionInfo.GetInfo_Between(findString1)[0];

        }

        /// <summary>
        /// 通用方法，通过开始字符串获取一个以字符串开始位置，此字符串右侧的的数据
        /// </summary>
        public string getSpecifyString_NextColumn(DataTable dt, string findString, int[] underThisPoints, int[] overThisPoints = null)
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
            //只获取第一次找到的字符串位置
            int[] lastFindPoints = new int[] { findString_Points[0][0], findString_Points[0][1] + 1 };
            //theDestString.Add(findStringPoint.GetInfo_Between(lastFindPoints));
            return findStringPoint.GetInfo_Between(lastFindPoints);
        }
        #endregion
    }

}
