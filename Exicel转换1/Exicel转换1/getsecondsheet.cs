﻿using System;
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
    class getsecondsheet
    {
        private string JOBName;
        private string Conveyor;
        private ExcelHelper excelHelper = null;
        private string sheetName_CustomerReport;
        private DataTable secondSheetDT = null;
        private DataTable smallsecondSheetDT = null;

        private int machineCount;
        private string sheetName_Result;
        private string Job;
        private string targetConveyor;

        public DataTable SecondSheetDT
        {
            get
            {
                return this.secondSheetDT;
            }
        }
        public DataTable SmallSecondSheetDT
        {
            get
            {
                return this.smallsecondSheetDT;
            }
        }

        //getsecondsheet构造函数
        public getsecondsheet(ExcelHelper excelHelper, string sheetName_CustomerReport, string JOBName, string Conveyor, int machineCount)
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

        #region 普通通用方法，获取DataTable中某一string的位置
        //获取DataTable中某一string的位置
        public int[] getPoints(DataTable dt, string findString)
        {

            getDataTablePointsInfo sectionInfo = new getDataTablePointsInfo(dt);
            int[][] Pannel_Points = sectionInfo.GetInfo_Between(findString);

            return Pannel_Points[0];
        }
        #endregion


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
            int[] lastFindPoints = new int[] { findString_Points[0][0], findString_Points[0][1] + 1 };
            //theDestString.Add(findStringPoint.GetInfo_Between(lastFindPoints));
            return findStringPoint.GetInfo_Between(lastFindPoints);
        }
    }
}
