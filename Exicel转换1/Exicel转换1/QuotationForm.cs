using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Exicel转换1
{
    public partial class QuotationForm : Form
    {
        #region 字段
        //evaluationReportClass
        EvaluationReportClass evaluationReportClass = null;
        //价格的基准记录
        List<double> unitPriceStandardList = null;
        //FeederNozzle的数量的基准记录。List indx和quotationList对应，int[]indx和行数对应，空的值为-1
        //List<int[]> NozzleQtyIntsList = null;
        //List<int> NozzleQtyList = null;
        //List<int[]> FeederQtyIntsList = null;
        //List<int> FeederQtyList = null;

        //最后生成quotationExcel的DataTable
        DataTable quotationExcelDT = null;

        //最低配置的总价格
        double minPrice = 0;
        public double MinPrice { get { return minPrice; }  }

        //QuotationDGVDTStruct结构
        QuotationDGVDTStruct quotationDGVDTStruct;

        //获取中间缺失的模组
        StringBuilder lackModuleBuilder = new StringBuilder();
        //检查是否缺少没有报价信息的Nozzle
        StringBuilder nozzleLackBuilder = new StringBuilder();
        //检查是否缺少没有报价信息的Feeder
        StringBuilder FeederLackBuilder = new StringBuilder();

        //数据结构，存放baseQotationDT、optionQotationDGVDT、moduleHeadQotationDGVDT、
        //nozzleQotationDGVDT、feederQotationDGVDT的开始和结束的在DataGridView中的行的索引位置
        struct DTPoint_Struct
        {
            public int[] baseQuotationDTStartEndRow;
            public List<int[]> ModuleQuotationDTStartEndRowList;
            public int[] OptionQuotationDTStartEndRow;
            public List<int[]> NozzleQuotationDTStartEndRowList;
            public List<int[]> FeederQuotationDTStartEndRowList;
        }

        //存放当前DataGridView中各个数据对应的DT的位置
        DTPoint_Struct dTPoint_Struct;
        #endregion

        
        public QuotationForm(EvaluationReportClass evaluationReportClass)
        {
            InitializeComponent();
            this.evaluationReportClass = evaluationReportClass;

            #region //DT和 DGV 结构的初始化
            quotationDGVDTStruct = new QuotationDGVDTStruct();
            //BaseQotation的 DataTable
            quotationDGVDTStruct.baseQotationDGVDT = new DataTable();
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("QTY");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("unit");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("DESCRIPTION");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add();
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("UNIT PRICE");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("AMOUNT");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("EP");
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("Option");//用于判定当前行是否可选
            quotationDGVDTStruct.baseQotationDGVDT.Columns.Add("QuatationType");//所属报价单的类型

            //OptionDataGridView的 DataTable
            quotationDGVDTStruct.optionQotationDGVDT = quotationDGVDTStruct.baseQotationDGVDT.Clone();

            //ModuleHeadDataGridView的 DataTable
            quotationDGVDTStruct.moduleHeadQotationDGVDT = quotationDGVDTStruct.baseQotationDGVDT.Clone();
            //添加QuatationType列，用于标识M6III/M3III

            quotationDGVDTStruct.nozzleQotationDGVDTList = new List<DataTable>();
            quotationDGVDTStruct.feederQotationDGVDTList = new List<DataTable>();

            //初始化quotationExcelDT
            quotationExcelDT = new DataTable();
            quotationExcelDT.Columns.Add("QTY", Type.GetType("System.Int32"));
            quotationExcelDT.Columns.Add("unit");
            quotationExcelDT.Columns.Add("DESCRIPTION");
            quotationExcelDT.Columns.Add();
            quotationExcelDT.Columns.Add("UNIT PRICE", Type.GetType("System.Int32"));
            quotationExcelDT.Columns.Add("AMOUNT", Type.GetType("System.Int32"));
            quotationExcelDT.Columns.Add("EP", Type.GetType("System.Int32"));

            //初始化QuotationDGV
            InitialQuotationDGV(); 
            #endregion

            #region //初始化quotationDT并显示最初的信息
            //获取所有的quotationDT
            AllQuotationDT allQuotationDT = AllQuotationDT.GetSingleAllQuotationDT();
            if (allQuotationDT == null)//获取失败
            {
                return;
            }

            //显示title标题版本 QuotationForm Text
            Text = "QuotationForm -" + allQuotationDT.Version;

            #region //根据当前job信息，生成对应的可供编辑的报价单信息

            bool needAdd;//需要添加到DataGridView中的标识，找不到数据则无法添加，标识是否添加成功
                         //QotationDataGV.EditMode= DataGridViewEditMode.
                         //QotationDataGV.allowus

            //base的报价单表
            int int_4MBASE = evaluationReportClass.base_StatisticsDict["4MBASE"];
            int int_2MBASE = evaluationReportClass.base_StatisticsDict["2MBASE"];

            #region //base及base的部品的BaseQotationDGVDT
            if (int_4MBASE != 0)
            {
                GetBaseAndBasePartsQotationDT(quotationDGVDTStruct.baseQotationDGVDT, allQuotationDT.BaseQotationDT,
                    int_4MBASE, "4MIII", allQuotationDT.PartsQotationDT, allQuotationDT.PartsQtyDT);
                //DataRow dr = BaseQotationDGV.NewRow();
                //dr["QTY"] = int_4MBASE;
                //dr[1] = "Set";
                //dr["DESCRIPTION"] = row_4MBaseQotation.FirstOrDefault().Field<string>("Machine Base");
                //dr["UNIT PRICE"] = row_4MBaseQotation.FirstOrDefault().Field<string>("JPYPrice");
                //dr["AMOUNT"] = Convert.ToInt32(row_4MBaseQotation.FirstOrDefault().Field<string>("JPYPrice")) * int_4MBASE;
                //dr["EP"] = Convert.ToInt32(row_4MBaseQotation.FirstOrDefault().Field<string>("EPPrice")) * int_4MBASE;
                //dr["Option"] = 0;//0为必选
                //BaseQotationDGV.Rows.Add(dr);

                //var rows_4MPartsQotation = from PartsQotation in PartsQotationDT.AsEnumerable()
                //                           join PartQty in PartsQtyDT.AsEnumerable()
                //                           on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                //                           where PartQty.Field<string>("basetype") == "4MIII"
                //                           select new
                //                           {
                //                               Qty = PartQty.Field<string>("num"),
                //                               Unit = PartQty.Field<string>("unit"),
                //                               PartsName = PartsQotation.Field<string>("PartsName"),
                //                               UnitPrice = PartsQotation.Field<string>("JPYPrice"),
                //                               AMOUNT = Convert.ToInt32(PartsQotation.Field<string>("JPYPrice")) * Convert.ToInt32(PartQty.Field<string>("num")),
                //                               Option = PartsQotation.Field<string>("option")
                //                           };
                //foreach (var item in rows_4MPartsQotation)
                //{
                //    //必选                                
                //        dr = BaseQotationDGV.NewRow();
                //        dr["QTY"] = item.Qty;
                //        dr[1] = item.Unit;
                //        dr["DESCRIPTION"] = item.PartsName;
                //        dr["UNIT PRICE"] = item.UnitPrice;
                //        dr["AMOUNT"] = item.AMOUNT;
                //        dr["Option"] = item.Option;
                //        BaseQotationDGV.Rows.Add(dr);                                 
                //}
            }
            if (int_2MBASE != 0)
            {
                GetBaseAndBasePartsQotationDT(quotationDGVDTStruct.baseQotationDGVDT, allQuotationDT.BaseQotationDT,
                    int_2MBASE, "2MIII", allQuotationDT.PartsQotationDT, allQuotationDT.PartsQtyDT);

            }
            #endregion

            #region //获取隔板数,和隔板的行
            int M3_Count = 0, M6_Count = 0;
            foreach (var item in evaluationReportClass.module_StatisticsDict)
            {

                if (item.Key.Substring(0, 2) == "M3")
                {
                    M3_Count += item.Value;
                }
                if (item.Key.Substring(0, 2) == "M6")
                {
                    M6_Count += item.Value;
                }
            }
            //隔板数
            int IntermediateCover_Count = M3_Count / 2 + M6_Count - 1;
            //获取隔板的DataTable
            var row_IntermediateCover = from PartQty in allQuotationDT.PartsQtyDT.AsEnumerable()
                                        join PartsQotation in allQuotationDT.PartsQotationDT.AsEnumerable()
                                        on PartQty.Field<string>("PartsName") equals PartsQotation.Field<string>("PartsName")
                                        where PartQty.Field<Int64>("identity") == 11
                                        select new
                                        {
                                            Qty = IntermediateCover_Count,
                                            Unit = PartQty.Field<string>("unit"),
                                            PartsName = PartsQotation.Field<string>("PartsName"),
                                            UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                            AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * IntermediateCover_Count,
                                            Option = PartQty.Field<Int64>("option")
                                        };

            DataRow dr;
            foreach (var item in row_IntermediateCover)
            {
                dr = quotationDGVDTStruct.optionQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Option :";
                quotationDGVDTStruct.optionQotationDGVDT.Rows.Add(dr);
            }
            #endregion

            #region //左右门
            var rows_LR_sidecover = from PartsQotation in allQuotationDT.PartsQotationDT.AsEnumerable()
                                    join PartQty in allQuotationDT.PartsQtyDT.AsEnumerable()
                                    on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                                    where (PartQty.Field<Int64>("identity") == 7 || PartQty.Field<Int64>("identity") == 8)
                                    select new
                                    {
                                        Qty = PartQty.Field<Int64>("num"),
                                        Unit = PartQty.Field<string>("unit"),
                                        PartsName = PartsQotation.Field<string>("PartsName"),
                                        UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                        AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * PartQty.Field<Int64>("num"),
                                        Option = PartQty.Field<Int64>("option")
                                    };

            //DataRow dr;
            foreach (var item in rows_LR_sidecover)
            {
                dr = quotationDGVDTStruct.optionQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Option :";
                quotationDGVDTStruct.optionQotationDGVDT.Rows.Add(dr);
            }
            #endregion

            #region //SMEMA线 
            //1m线
            var rows_SMEMACable1m = from PartsQotation in allQuotationDT.PartsQotationDT.AsEnumerable()
                                    join PartQty in allQuotationDT.PartsQtyDT.AsEnumerable()
                                    on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                                    where PartQty.Field<Int64>("identity") == 9
                                    select new
                                    {
                                        Qty = (int_4MBASE + int_2MBASE - 1) * 2,
                                        Unit = PartQty.Field<string>("unit"),
                                        PartsName = PartsQotation.Field<string>("PartsName"),
                                        UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                        AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * PartQty.Field<Int64>("num"),
                                        Option = PartQty.Field<Int64>("option")
                                    };

            //DataRow dr;
            foreach (var item in rows_SMEMACable1m)
            {
                dr = quotationDGVDTStruct.optionQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Option :";
                quotationDGVDTStruct.optionQotationDGVDT.Rows.Add(dr);
            }
            //1.5 m线
            var rows_SMEMACable1_5m = from PartsQotation in allQuotationDT.PartsQotationDT.AsEnumerable()
                                      join PartQty in allQuotationDT.PartsQtyDT.AsEnumerable()
                                      on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                                      where PartQty.Field<Int64>("identity") == 10
                                      select new
                                      {
                                          Qty = 4,
                                          Unit = PartQty.Field<string>("unit"),
                                          PartsName = PartsQotation.Field<string>("PartsName"),
                                          UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                          AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * PartQty.Field<Int64>("num"),
                                          Option = PartQty.Field<Int64>("option")
                                      };

            //DataRow dr;
            foreach (var item in rows_SMEMACable1_5m)
            {
                dr = quotationDGVDTStruct.optionQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Option :";
                quotationDGVDTStruct.optionQotationDGVDT.Rows.Add(dr);
            }
            #endregion

            #region //上料站Reel setting stand  

            var rows_Stand = from PartsQotation in allQuotationDT.PartsQotationDT.AsEnumerable()
                             join PartQty in allQuotationDT.PartsQtyDT.AsEnumerable()
                             on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                             where PartQty.Field<string>("AccessoryType") == "Accessories"
                             select new
                             {
                                 Qty = (int_4MBASE + int_2MBASE - 1) * 2,
                                 Unit = PartQty.Field<string>("unit"),
                                 PartsName = PartsQotation.Field<string>("PartsName"),
                                 UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                 AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * PartQty.Field<Int64>("num"),
                                 Option = PartQty.Field<Int64>("option")
                             };

            //DataRow dr;
            foreach (var item in rows_Stand)
            {
                dr = quotationDGVDTStruct.optionQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Option :";
                quotationDGVDTStruct.optionQotationDGVDT.Rows.Add(dr);
            }

            #endregion

            #region //ModuleHeadDGVDT
            //ModuleHeadQotationDT
            //统计module head的个数(module head均相同，算作一个)
            Dictionary<string, int> moduleHeadStatisticsDict = new Dictionary<string, int>();
            for (int i = 0; i < evaluationReportClass.AllModuleType.Length; i++)
            {
                //拼接module head作为Key,","
                string key = string.Join(",", evaluationReportClass.AllModuleType[i], evaluationReportClass.AllHeadType[i]);
                if (moduleHeadStatisticsDict.ContainsKey(key))
                {
                    moduleHeadStatisticsDict[key] += 1;
                }
                else
                {
                    moduleHeadStatisticsDict.Add(key, 1);
                }
            }

            //统计MTray个数，最后统一获取MTray
            int MTrayCount = 0;
            //获取ModuleHeadQotationDT
            foreach (var moduleHeadStatistics in moduleHeadStatisticsDict)
            {
                //IEnumerable<> rows_ModuleHeadQotation;
                string[] moduleHeadType = moduleHeadStatistics.Key.Split(',');
                string[] modules = moduleHeadType[0].Split('-');
                if (modules.Length > 1)
                {
                    //由于懒加载的机制，需要设定一个flag，是否查找添加到对应的DataTable
                    needAdd = true;
                    if (modules[1]=="MTray")//Mtray时，作为Feeder处理
                    {
                        MTrayCount++;
                        #region //获取模组
                        var rows_ModuleHeadQotation = from ModuleHeadQotation in allQuotationDT.ModuleHeadQotationDT.AsEnumerable()
                                                      where (ModuleHeadQotation.Field<string>("ModuleType") == moduleHeadType[0] &&
                                                      string.IsNullOrEmpty(ModuleHeadQotation.Field<string>("TrayType")) &&
                                                      ModuleHeadQotation.Field<string>("HeadType") == moduleHeadType[1] &&
                                                      ModuleHeadQotation.Field<string>("ProductionMode") == (evaluationReportClass.ProductionMode.Split('-')[0] == "Dual" ? "Double" : "Single"))
                                                      select new
                                                      {
                                                          Qty = moduleHeadStatistics.Value,
                                                          Unit = moduleHeadStatistics.Value > 1 ? "Sets" : "Set",
                                                          description = ModuleHeadQotation.Field<string>("Includes"),
                                                          UnitPrice = ModuleHeadQotation.Field<Int64>("JPYPrice"),
                                                          AMOUNT = ModuleHeadQotation.Field<Int64>("JPYPrice") * moduleHeadStatistics.Value,
                                                          Option = ModuleHeadQotation.Field<Int64>("Option"),
                                                          moduleType = ModuleHeadQotation.Field<string>("ModuleType") + " Module " + ModuleHeadQotation.Field<string>("PartCamera") + " :"
                                                      };
                        //DataRow dr;
                        foreach (var item in rows_ModuleHeadQotation)
                        {
                            dr = quotationDGVDTStruct.moduleHeadQotationDGVDT.NewRow();
                            dr["QTY"] = item.Qty;
                            dr[1] = item.Unit;
                            dr["DESCRIPTION"] = item.description;
                            dr["UNIT PRICE"] = item.UnitPrice;
                            dr["AMOUNT"] = item.AMOUNT;
                            dr["Option"] = item.Option;
                            dr["QuatationType"] = item.moduleType;
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows.Add(dr);
                            needAdd = false;
                        }
                        //if (needAdd)
                        //{
                        //    lackModuleBuilder.Append(moduleHeadType[0] + "\t" + moduleHeadType[1] + "\r\n");
                        //} 
                        #endregion
                        
                    }
                    else
                    {
                        var rows_ModuleHeadQotation = from ModuleHeadQotation in allQuotationDT.ModuleHeadQotationDT.AsEnumerable()
                                                      where (ModuleHeadQotation.Field<string>("ModuleType") == modules[0] &&
                                                      ModuleHeadQotation.Field<string>("TrayType") == modules[1].Substring(0, modules[1].Length - 4) &&
                                                      ModuleHeadQotation.Field<string>("HeadType") == moduleHeadType[1] &&
                                                      ModuleHeadQotation.Field<string>("ProductionMode") == (evaluationReportClass.ProductionMode.Split('-')[0] == "Dual" ? "Double" : "Single"))
                                                      select new
                                                      {
                                                          Qty = moduleHeadStatistics.Value,
                                                          Unit = moduleHeadStatistics.Value > 1 ? "Sets" : "Set",
                                                          description = ModuleHeadQotation.Field<string>("Includes"),
                                                          UnitPrice = ModuleHeadQotation.Field<Int64>("JPYPrice"),
                                                          AMOUNT = ModuleHeadQotation.Field<Int64>("JPYPrice") * moduleHeadStatistics.Value,
                                                          Option = ModuleHeadQotation.Field<Int64>("Option"),
                                                          moduleType = $"{ ModuleHeadQotation.Field<string>("ModuleType")} Module {ModuleHeadQotation.Field<string>("PartCamera")}:"
                                                      };
                        //DataRow dr;
                        foreach (var item in rows_ModuleHeadQotation)
                        {
                            dr = quotationDGVDTStruct.moduleHeadQotationDGVDT.NewRow();
                            dr["QTY"] = item.Qty;
                            dr[1] = item.Unit;
                            dr["DESCRIPTION"] = item.description;
                            dr["UNIT PRICE"] = item.UnitPrice;
                            dr["AMOUNT"] = item.AMOUNT;
                            dr["Option"] = item.Option;
                            dr["QuatationType"] = item.moduleType;
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows.Add(dr);
                            //已经添加
                            needAdd = false;
                        }                        
                    }
                    if (needAdd)
                    {
                        lackModuleBuilder.Append(moduleHeadType[0] + "\t" + moduleHeadType[1] + "\r\n");
                    }
                }
                else
                {
                    needAdd = true;
                    var rows_ModuleHeadQotation = from ModuleHeadQotation in allQuotationDT.ModuleHeadQotationDT.AsEnumerable()
                                                  where (ModuleHeadQotation.Field<string>("ModuleType") == moduleHeadType[0] &&
                                                  string.IsNullOrEmpty(ModuleHeadQotation.Field<string>("TrayType")) &&
                                                  ModuleHeadQotation.Field<string>("HeadType") == moduleHeadType[1] &&
                                                  ModuleHeadQotation.Field<string>("ProductionMode") == (evaluationReportClass.ProductionMode.Split('-')[0] == "Dual" ? "Double" : "Single"))
                                                  select new
                                                  {
                                                      Qty = moduleHeadStatistics.Value,
                                                      Unit = moduleHeadStatistics.Value > 1 ? "Sets" : "Set",
                                                      description = ModuleHeadQotation.Field<string>("Includes"),
                                                      UnitPrice = ModuleHeadQotation.Field<Int64>("JPYPrice"),
                                                      AMOUNT = ModuleHeadQotation.Field<Int64>("JPYPrice") * moduleHeadStatistics.Value,
                                                      Option = ModuleHeadQotation.Field<Int64>("Option"),
                                                      moduleType = ModuleHeadQotation.Field<string>("ModuleType") + " Module " + ModuleHeadQotation.Field<string>("PartCamera") + " :"
                                                  };
                    //DataRow dr;
                    foreach (var item in rows_ModuleHeadQotation)
                    {
                        dr = quotationDGVDTStruct.moduleHeadQotationDGVDT.NewRow();
                        dr["QTY"] = item.Qty;
                        dr[1] = item.Unit;
                        dr["DESCRIPTION"] = item.description;
                        dr["UNIT PRICE"] = item.UnitPrice;
                        dr["AMOUNT"] = item.AMOUNT;
                        dr["Option"] = item.Option;
                        dr["QuatationType"] = item.moduleType;
                        quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows.Add(dr);
                        needAdd = false;
                    }
                    if (needAdd)
                    {
                        lackModuleBuilder.Append(moduleHeadType[0] + "\t" + moduleHeadType[1] + "\r\n");
                    }
                }
            }
                      

            if (lackModuleBuilder.Length != 0)
            {
                //有缺少的数据
                return;//缺少模组直接退出，不生成
            }
            #endregion

            #region //NozzleQuotation
            ////////////计算head NozzleType 和 NozzleQty
            //Nozzle_Statistics吸嘴的统计字典
            Dictionary<string, int> Nozzle_StatisticsDict = new Dictionary<string, int>();
            for (int i = 0; i < evaluationReportClass.feederNozzleDT.Rows.Count; i++)
            {
                //空行时
                if (evaluationReportClass.feederNozzleDT.Rows[i]["Nozzle"] == null || evaluationReportClass.feederNozzleDT.Rows[i]["Nozzle"].Equals(DBNull.Value))
                {
                    continue;
                }
                Nozzle_StatisticsDict.Add(evaluationReportClass.feederNozzleDT.Rows[i]["Nozzle"].ToString(),
                Convert.ToInt32(evaluationReportClass.feederNozzleDT.Rows[i]["QTy Nozzle"]));
            }
            //重新整合以便Head，因为含有R12/H08/H12HS/V12，为防止写死(单独判断)硬编码，添加数据后需要改代码
            //因此首先整合一遍，使之对应于R12/H08/:Nozzle
            //for (int i = 0; i < evaluationReportClass.HeadTypeNozzleDict.Keys.Count; i++)
            //{
            //    from NozzleQotation in allQuotationDT.NozzleQotationDT.AsEnumerable()
            //    where (NozzleQotation.Field<string>("HeadType").Contains(evaluationReportClass.HeadTypeNozzleDict.Keys.)
            //}

            //evaluationReportClass.HeadTypeNozzleDict.Keys.Contains()
            //Dictionary<>
            //foreach (var item in evaluationReportClass.HeadTypeNozzleDict)
            //{

            //}
                        

            //以上实现较困难.借助于NozzleDetail唯一即Nozzle_name唯一
            //中转临时DataTable，合并nozzlename相同。借助于quotationDGVDTStruct的"EP"字段，不使用临时DataTable
            //转换为DataTableList的中转临时DataTable
            DataTable _tempDataTable = quotationDGVDTStruct.baseQotationDGVDT.Clone();
            //HeadtypeNzoole 获取quotation
            try
            {
                foreach (var HeadTypeNozzle in evaluationReportClass.HeadTypeNozzleDict)
                {
                    //判断是否和拆分
                    //if (HeadTypeNozzle.Value.Where())
                    //{

                    //}
                    //循环获取 headtype和Nozzle 对应的quotation数据
                    //H04SF H04这两种情况，如果为H04 Contains将会选择两个
                    //where (NozzleQuotation.Field<string>("HeadType").Contains(HeadTypeNozzle.Key) &&
                    var rows_NozzleQuotation = from NozzleQuotation in allQuotationDT.NozzleQotationDT.AsEnumerable()
                                               where (NozzleQuotation.Field<string>("HeadType").Split('/').Contains(HeadTypeNozzle.Key) &&
                                               (string.IsNullOrEmpty(NozzleQuotation.Field<string>("Nozzle")) ||
                                               HeadTypeNozzle.Value.FirstOrDefault(p => p.Split('-')[1] == NozzleQuotation.Field<string>("Nozzle")) != null))
                                               select new
                                               {
                                                   Unit = NozzleQuotation.Field<string>("unit"),
                                                   description = NozzleQuotation.Field<string>("NozzleDetial"),
                                                   UnitPrice = NozzleQuotation.Field<Int64>("JPYPrice"),
                                                   //AMOUNT = ModuleHeadQotation.Field<Int64>("JPYPrice") * moduleHeadStatistics.Value,
                                                   Option = string.IsNullOrEmpty(NozzleQuotation.Field<string>("Nozzle")) ? 1 : 0,
                                                   nozzleType = NozzleQuotation.Field<string>("NozzleTitle"),
                                                   nozzleName = HeadTypeNozzle.Value.FirstOrDefault(p => p.Split('-')[1] == NozzleQuotation.Field<string>("Nozzle"))
                                                   //nozzleName为null的判断
                                               };


                    //quotationDGVDTStruct
                    //遍历处理是否有相同的nozzlename
                    foreach (var item in rows_NozzleQuotation)
                    {
                        //空值时，即为storage或dumpyNozzle，也即option==1
                        if (item.nozzleName == null)
                        {//空值，将其行添加到nozzleQotationDGVDT
                            dr = _tempDataTable.NewRow();
                            dr["QTY"] = 0;
                            dr[1] = item.Unit;
                            dr["DESCRIPTION"] = item.description;
                            dr["UNIT PRICE"] = item.UnitPrice;
                            dr["AMOUNT"] = 0;
                            dr["Option"] = item.Option;
                            dr["QuatationType"] = item.nozzleType;
                            _tempDataTable.Rows.Add(dr);
                            continue;
                        }
                        //不为空时遍历，是否有相同
                        bool isHasEqual = false;
                        for (int i = 0; i < _tempDataTable.Rows.Count; i++)
                        {
                            //存在相同nozzlename
                            if (!_tempDataTable.Rows[i]["EP"].Equals(DBNull.Value) && _tempDataTable.Rows[i]["EP"].ToString() == item.nozzleName)
                            {
                                isHasEqual = true;
                                break;//存在，设置为存在，结束
                            }
                        }
                        //判断是否存在相同的nozzlename，不存在添加
                        if (!isHasEqual)//不存在相同nozzlename
                        {
                            dr = _tempDataTable.NewRow();
                            //dr["QTY"] = no;
                            dr[1] = item.Unit;
                            dr["DESCRIPTION"] = item.description;
                            dr["UNIT PRICE"] = item.UnitPrice;
                            //dr["AMOUNT"] = 0;
                            dr["Option"] = item.Option;
                            dr["QuatationType"] = item.nozzleType;
                            dr["EP"] = item.nozzleName;
                            _tempDataTable.Rows.Add(dr);
                        }
                    }
                }
                //处理结束后，给nozzleName的QTY的price赋值
                for (int i = 0; i < _tempDataTable.Rows.Count; i++)
                {
                    if (_tempDataTable.Rows[i]["EP"] != null && !_tempDataTable.Rows[i]["EP"].Equals(DBNull.Value))//不为空时
                    {
                        _tempDataTable.Rows[i]["QTY"] = Nozzle_StatisticsDict[_tempDataTable.Rows[i]["EP"].ToString()];
                        _tempDataTable.Rows[i]["AMOUNT"] = Nozzle_StatisticsDict[_tempDataTable.Rows[i]["EP"].ToString()] *
                            Convert.ToInt32(_tempDataTable.Rows[i]["UNIT PRICE"]);
                    }
                }

                //判断是否缺少Nozzle
                foreach (var item in Nozzle_StatisticsDict)
                {
                    bool isExists = false;//默认不存在
                    foreach (DataRow dataRow in _tempDataTable.Rows)
                    {
                        //如果存在
                        if (item.Key == dataRow["EP"].ToString())
                        {
                            isExists = true;
                            break;
                        }

                    }
                    //不存在
                    if (!isExists)
                    {
                        nozzleLackBuilder.Append(item.Key + "\t" + item.Value);
                    }
                }

                //临时变量按QuotationType分类，拆分为DataTableList.//linq的查询延续的写法不知怎么返回和接受多个group分组
                var _nozzleQuotationDT_Group = from tempRow in _tempDataTable.AsEnumerable()
                                               group tempRow by tempRow.Field<string>("QuatationType");

                foreach (var g in _nozzleQuotationDT_Group)
                {
                    ///g.Key;//分组键
                    DataTable a = _tempDataTable.Clone();
                    foreach (var item in g)
                    {
                        DataRow adr = a.NewRow();
                        adr["QTY"] = item.Field<string>("QTY");
                        adr[1] = item.Field<string>(1);
                        adr["DESCRIPTION"] = item.Field<string>("DESCRIPTION");
                        adr["UNIT PRICE"] = item.Field<string>("UNIT PRICE");
                        adr["AMOUNT"] = item.Field<string>("AMOUNT");
                        adr["Option"] = item.Field<string>("Option");
                        adr["QuatationType"] = item.Field<string>("QuatationType"); ;
                        //adr["EP"] = item.nozzleName;
                        a.Rows.Add(adr);
                    }
                    quotationDGVDTStruct.nozzleQotationDGVDTList.Add(a);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #endregion

            #region //feederQuotation
            //使用_tempDataTable临时变量，清空数据
            _tempDataTable.Clear();
            //不中转获取feeder_statistics
            //Dictionary<string, int> feeder_Statistics = new Dictionary<string, int>();            
            try
            {
                #region //有MTray时，将MTray加入FeederQuotation中
                if (MTrayCount != 0)
                {
                    #region //获取Mtray
                    var rows_FeederQuotation = from feederQuotation in allQuotationDT.FeederQotationDT.AsEnumerable()
                                               where feederQuotation.Field<string>("Feeder") == "MTray"
                                               select new
                                               {
                                                   Qty = MTrayCount,
                                                   Unit = feederQuotation.Field<string>("unit"),
                                                   description = feederQuotation.Field<string>("FeederDetail"),
                                                   UnitPrice = feederQuotation.Field<Int64>("JPYPrice"),
                                                   AMOUNT = feederQuotation.Field<Int64>("JPYPrice") * MTrayCount,
                                                   Option = feederQuotation.Field<Int64>("Option"),
                                                   feederType = feederQuotation.Field<string>("FeederTitle")
                                               };

                    foreach (var item in rows_FeederQuotation)
                    {
                        dr = _tempDataTable.NewRow();
                        dr["QTY"] = item.Qty;
                        dr[1] = item.Unit;
                        dr["DESCRIPTION"] = item.description;
                        dr["UNIT PRICE"] = item.UnitPrice;
                        dr["AMOUNT"] = item.AMOUNT;
                        dr["Option"] = item.Option;
                        dr["QuatationType"] = item.feederType;
                        _tempDataTable.Rows.Add(dr);
                    }
                    #endregion
                }
                #endregion

                for (int i = 0; i < evaluationReportClass.feederNozzleDT.Rows.Count; i++)
                {
                    //索引超出了数组界限
                    if (evaluationReportClass.feederNozzleDT.Rows[i]["Feeder"] == null || evaluationReportClass.feederNozzleDT.Rows[i]["Feeder"].Equals(DBNull.Value))
                    {
                        continue;
                    }
                    string feeder = evaluationReportClass.feederNozzleDT.Rows[i]["Feeder"].ToString();
                    if (feeder.Split('-').Length < 3)
                    {
                        //不符合标准命名的feeder，quotation中没有，且索引取值会报错
                        continue;
                    }
                    //特殊吸嘴 是否可以按照-划分
                    var rows_FeederQuotation = from feederQuotation in allQuotationDT.FeederQotationDT.AsEnumerable()
                                               where feederQuotation.Field<string>("Feeder").Contains(feeder.Split('-')[1]) &&
                                               (feederQuotation.Field<string>("Size") == feeder.Split('-')[2] ||
                                               feederQuotation.Field<string>("Size") == "0")
                                               select new
                                               {
                                                   Qty = Convert.ToInt32(evaluationReportClass.feederNozzleDT.Rows[i]["Qty Feeder"]),
                                                   Unit = feederQuotation.Field<string>("unit"),
                                                   description = feederQuotation.Field<string>("FeederDetail"),
                                                   UnitPrice = feederQuotation.Field<Int64>("JPYPrice"),
                                                   AMOUNT = feederQuotation.Field<Int64>("JPYPrice") * Convert.ToInt32(evaluationReportClass.feederNozzleDT.Rows[i]["Qty Feeder"]),
                                                   Option = feederQuotation.Field<Int64>("Option"),
                                                   feederType = feederQuotation.Field<string>("FeederTitle"),
                                                   feederName = feeder
                                               };

                    foreach (var item in rows_FeederQuotation)
                    {
                        dr = _tempDataTable.NewRow();
                        dr["QTY"] = item.Qty;
                        dr[1] = item.Unit;
                        dr["DESCRIPTION"] = item.description;
                        dr["UNIT PRICE"] = item.UnitPrice;
                        dr["AMOUNT"] = item.AMOUNT;
                        dr["EP"] = item.feederName;
                        dr["Option"] = item.Option;
                        dr["QuatationType"] = item.feederType;
                        _tempDataTable.Rows.Add(dr);
                    }
                }
                //判断是否缺少Feeder
                foreach (DataRow itemDR in evaluationReportClass.feederNozzleDT.Rows)
                {
                    //去掉空的cell
                    if (itemDR["Feeder"] == null || itemDR["Feeder"].Equals(DBNull.Value))
                    {
                        continue;
                    }

                    bool isExists = false;//默认不存在
                    foreach (DataRow dataRow in _tempDataTable.Rows)
                    {
                        //如果存在
                        if (!dataRow["EP"].Equals(DBNull.Value)&&itemDR["Feeder"].ToString() == dataRow["EP"].ToString())
                        {
                            isExists = true;
                            break;
                        }
                    }
                    //不存在
                    if (!isExists)
                    {
                        FeederLackBuilder.Append(itemDR["Feeder"].ToString() + "\t" + itemDR["Qty Feeder"].ToString());
                    }
                }

                //临时变量按Quotation分组，放到feederQotationDGVDTList中
                var _feederQuotationDT_Group = from tempRow in _tempDataTable.AsEnumerable()
                                               group tempRow by tempRow.Field<string>("QuatationType");

                foreach (var g in _feederQuotationDT_Group)
                {
                    ///g.Key;//分组键
                    DataTable a = _tempDataTable.Clone();
                    foreach (var item in g)
                    {
                        DataRow adr = a.NewRow();
                        adr["QTY"] = item.Field<string>("QTY");
                        adr[1] = item.Field<string>(1);
                        adr["DESCRIPTION"] = item.Field<string>("DESCRIPTION");
                        adr["UNIT PRICE"] = item.Field<string>("UNIT PRICE");
                        adr["AMOUNT"] = item.Field<string>("AMOUNT");
                        adr["Option"] = item.Field<string>("Option");
                        adr["QuatationType"] = item.Field<string>("QuatationType"); ;
                        //adr["EP"] = item.nozzleName;
                        a.Rows.Add(adr);
                    }
                    quotationDGVDTStruct.feederQotationDGVDTList.Add(a);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            

            #endregion 
            #endregion

            //QuotationDGVDT中显示各报价单 及最后最低的价格之和
            ShowQuotationDGVDT(quotationDGVDTStruct);
            #endregion
        }

        // 重写OnClosing使点击关闭按键时隐藏窗体
        protected override void OnClosing(CancelEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }

        #region //GetBaseAndBasePartsQotationDT的方法，获取base相关的quotation
        private void GetBaseAndBasePartsQotationDT(DataTable BaseQotationDGVDT, DataTable BaseQotationDT, int int_BASE,
            string baseName, DataTable PartsQotationDT, DataTable PartsQtyDT)
        {
            //base
            var row_4MBaseQotation = from dr1 in BaseQotationDT.AsEnumerable()
                                     where dr1.Field<string>("BaseType") == baseName
                                     select dr1;
            //var row_2MBaseQotation = from dr in BaseQotationDT.AsEnumerable()
            //                         where dr.Field<string>("BaseType") == "2MIII"
            //                         select dr;

            DataRow dr = BaseQotationDGVDT.NewRow();
            dr["QTY"] = int_BASE;
            dr[1] = "Set";
            dr["DESCRIPTION"] = row_4MBaseQotation.FirstOrDefault().Field<string>("Machine Base");
            dr["UNIT PRICE"] = row_4MBaseQotation.FirstOrDefault().Field<Int64>("JPYPrice");
            dr["AMOUNT"] = row_4MBaseQotation.FirstOrDefault().Field<Int64>("JPYPrice") * int_BASE;
            dr["EP"] = row_4MBaseQotation.FirstOrDefault().Field<Int64>("EPPrice") * int_BASE;
            dr["Option"] = 0;//0为必选
            dr["QuatationType"] = "Machine Base :";
            BaseQotationDGVDT.Rows.Add(dr);

            //与base对应的Parts
            var rows_4MPartsQotation = from PartsQotation in PartsQotationDT.AsEnumerable()
                                       join PartQty in PartsQtyDT.AsEnumerable()
                                       on PartsQotation.Field<string>("PartsName") equals PartQty.Field<string>("PartsName")
                                       where PartQty.Field<string>("basetype") == baseName
                                       select new
                                       {
                                           Qty = PartQty.Field<Int64>("num") * int_BASE,
                                           Unit = PartQty.Field<string>("unit"),
                                           PartsName = PartsQotation.Field<string>("PartsName"),
                                           UnitPrice = PartsQotation.Field<Int64>("JPYPrice"),
                                           AMOUNT = PartsQotation.Field<Int64>("JPYPrice") * PartQty.Field<Int64>("num") * int_BASE,
                                           Option = PartQty.Field<Int64>("option")
                                       };
            foreach (var item in rows_4MPartsQotation)
            {
                //必选                                
                dr = BaseQotationDGVDT.NewRow();
                dr["QTY"] = item.Qty;
                dr[1] = item.Unit;
                dr["DESCRIPTION"] = item.PartsName;
                dr["UNIT PRICE"] = item.UnitPrice;
                dr["AMOUNT"] = item.AMOUNT;
                dr["Option"] = item.Option;
                dr["QuatationType"] = "Machine Base :";
                BaseQotationDGVDT.Rows.Add(dr);
            }
        } 
        #endregion
        
        //显示窗体的方法，用以在每次显示时，对于缺少的报价单进行显示
        public void ShowForm()
        {
            #region //判断缺少的报价单并显示。
            //判断是否可以生成报价单
            if (lackModuleBuilder.Length != 0)
            {
                //有缺少的数据
                MessageBox.Show("无法生成报价单！请确认");
                QuotationLackItem quotationLack = new QuotationLackItem("缺少以下Module \t Head的报价单信息：\r\n\r\n" + lackModuleBuilder.ToString());
                if (quotationLack.ShowDialog() == DialogResult.OK)
                {
                    //销毁资源
                    Dispose();
                    return;
                }
            }
            //缺少的Nozzle和feeder
            if (nozzleLackBuilder.Length != 0)
            {
                nozzleLackBuilder.Insert(0, "NozzleName\t Qty");
                if (FeederLackBuilder.Length != 0)
                {
                    FeederLackBuilder.Insert(0, "NozzleName\t Qty");
                }

                QuotationLackItem quotationLack = new QuotationLackItem(
                    nozzleLackBuilder.ToString() + "\r\n\r\n" + FeederLackBuilder.ToString());
                quotationLack.ShowDialog();
            }
            #endregion
            //显示
            Show();
        }

        private void QotationForm_Load(object sender, EventArgs e)
        {
            #region //初始状态
            //折扣按钮默认不选中
            discountCheckB.Checked = false;
            discountComboB.Enabled = false;
            label1.Enabled = false;
            //默认显示10折
            discountComboB.SelectedItem = discountComboB.Items[0];

            //feederNozzle数量调整功能部分的默认状态和数据
            foreach (Control control in feederNozzlegroupB.Controls)
            {
                CheckBox checkBox = control as CheckBox;
                ComboBox comboBox = control as ComboBox;
                if (comboBox != null)
                {
                    comboBox.SelectedItem = comboBox.Items[0];
                }
                if (checkBox != null)
                {
                    checkBox.Checked = false;
                }
                else
                {
                    control.Enabled = false;
                }
            }
            #endregion

            

            //注册事件放在初始化完成之后，否则在初始化时就触发事件
            //添加事件
            discountComboB.Leave += discountComboB_TextChanged;            
            //discountComboB_TextChanged
            QuotationDataGV.CellValueChanged += QotationDataGV_CellValueChanged;
            QuotationDataGV.CellContentClick += QotationDataGV_CellContentClick;
            //QotationDataGV.CellContentClick += QotationDataGV_CellContentClick;
                        
            FeederQtyComboB.TextChanged += FeederQtyComboB_TextChanged;            
            NozzleQtyComboB.TextChanged += NozzleQtyComboB_TextChanged;
        }

        private void QotationDataGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == null)
            {
                return;
            }
            if (e.ColumnIndex == 0&& e.RowIndex< dataGridView.Rows.Count-3)//选中和取消当前
            {
                
                //当前行的 取值 click点击后就相反
                bool currentChecked = (bool)((DataGridViewCheckBoxCell)dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex]).EditedFormattedValue;
                //((DataGridViewCheckBoxCell)dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex]).Value = currentChecked;

                //判断修改的是哪一DataTable中的数据
                //先判定在属于哪个DT
                if (e.RowIndex >= dTPoint_Struct.baseQuotationDTStartEndRow[0] && e.RowIndex < dTPoint_Struct.baseQuotationDTStartEndRow[1])
                {
                    //选中和取消 对应的索引
                    int indx = e.RowIndex - dTPoint_Struct.baseQuotationDTStartEndRow[0];
                    quotationDGVDTStruct.baseQotationDGVDT.Rows[indx]["Option"] = currentChecked ? 0 : 1;
                }
                else if (e.RowIndex >= dTPoint_Struct.ModuleQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.ModuleQuotationDTStartEndRowList[dTPoint_Struct.ModuleQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.ModuleQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前module报价单块，开始第一行即为选中的行。即当前第i个（moduleHeadQotationDGVDT的第i行）选中
                        if (dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] == e.RowIndex)
                        {
                            //int indx = e.RowIndex - dTPoint_Struct.baseQotationDTStartEndRow[0];
                            int a = currentChecked ? 0 : 1;
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["Option"] = a;
                        }
                    }
                }
                else if (e.RowIndex >= dTPoint_Struct.OptionQuotationDTStartEndRow[0] && e.RowIndex < dTPoint_Struct.OptionQuotationDTStartEndRow[1])
                {
                    //选中和取消 对应的索引-1,,有标题行
                    int indx = e.RowIndex - dTPoint_Struct.OptionQuotationDTStartEndRow[0] - 1;
                    quotationDGVDTStruct.optionQotationDGVDT.Rows[indx]["Option"] = currentChecked ? 0 : 1;
                }
                else if (e.RowIndex >= dTPoint_Struct.NozzleQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.NozzleQuotationDTStartEndRowList[dTPoint_Struct.NozzleQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.NozzleQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前nozzle报价单块，选中的行。为当前QuatationType下的第 indx = e.RowIndex - dTPoint_Struct.FeederDTStartEndRow[0]-1索引
                        //多一行标题
                        //i 即为第i个QuatationType，即按QuatationType分组后的DataTableList的indx
                        if (e.RowIndex >= dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] &&
                            e.RowIndex < dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][1])
                        {
                            int indx = e.RowIndex - dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] - 1;
                            quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[indx]["Option"] = currentChecked ? 0 : 1;
                        }
                    }
                }
                else if (e.RowIndex >= dTPoint_Struct.FeederQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.FeederQuotationDTStartEndRowList[dTPoint_Struct.FeederQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.FeederQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前nozzle报价单块，选中的行。为当前QuatationType下的第 indx = e.RowIndex - dTPoint_Struct.FeederDTStartEndRow[0]-1索引
                        //多一行标题
                        //i 即为第i个QuatationType，即按QuatationType分组后的DataTableList的indx
                        if (e.RowIndex >= dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] &&
                            e.RowIndex < dTPoint_Struct.FeederQuotationDTStartEndRowList[i][1])
                        {
                            int indx = e.RowIndex - dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] - 1;
                            quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[indx]["Option"] = currentChecked ? 0 : 1;
                        }
                    }
                }
                //重新计算sum
                QuotationDataGVPriceSum(QuotationDataGV);
            }

        }

        private void discountCheckB_CheckedChanged(object sender, EventArgs e)
        {
            if (discountCheckB.Checked)
            {
                label1.Enabled = true;
                discountComboB.Enabled = true;
            }
            else
            {
                label1.Enabled = false;
                discountComboB.Enabled = false;
                //取消折扣选择时,变为默认的10折扣，同时价格复位
                double disc = (double)10 / Convert.ToDouble(discountComboB.Text);
                discountComboB.SelectedItem = discountComboB.Items[0];
                //对单价还原
                DisCountUnitPrice(disc);
            }
        }

        #region //初始化DataGridView
        void InitialQuotationDGV()
        {
            //实现DataGridView
            //DataGridView禁止排序
            QuotationDataGV.AllowUserToAddRows = false;
            QuotationDataGV.AllowUserToDeleteRows = false;
            QuotationDataGV.AllowUserToOrderColumns = false;
            //QotationDataGV.AllowUserToResizeColumns = false;
            QuotationDataGV.AllowUserToResizeRows = false;

            //允许编辑
            QuotationDataGV.ReadOnly = false;
            //DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
            //checkBoxColumn.HeaderText = "选择";
            //QotationDataGV.Columns.Add(checkBoxColumn);
            QuotationDataGV.Columns.Add("Select", "选择");
            //添加其他列
            QuotationDataGV.Columns.Add("QTY", "QTY");
            QuotationDataGV.Columns.Add("unit", "");
            QuotationDataGV.Columns.Add("DESCRIPTION", "DESCRIPTION");
            QuotationDataGV.Columns.Add("", "");
            QuotationDataGV.Columns.Add("UNIT PRICE", "UNIT PRICE");
            QuotationDataGV.Columns.Add("AMOUNT", "AMOUNT");
            QuotationDataGV.Columns.Add("EP", "EP");

            //样式居中
            QuotationDataGV.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            QuotationDataGV.Columns["DESCRIPTION"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //禁用列的自动排序                        
            for (int i = 0; i < QuotationDataGV.Columns.Count; i++)
            {
                QuotationDataGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                //禁用编辑，除了Price列、EP列、和数量列(1)
                if (i == 0 || i == 1 || QuotationDataGV.Columns[i].Name == "UNIT PRICE" || QuotationDataGV.Columns[i].Name == "AMOUNT" ||
                    QuotationDataGV.Columns[i].Name == "EP")
                {
                    QuotationDataGV.Columns[i].ReadOnly = false;
                }
                else
                {
                    QuotationDataGV.Columns[i].ReadOnly = true;
                }
            }
            //禁用行标题
            QuotationDataGV.RowHeadersVisible = false;
            //设置列的宽度
            QuotationDataGV.Columns[0].Width = 34;
            QuotationDataGV.Columns[1].Width = 90;
            QuotationDataGV.Columns[2].Width = 40;
            QuotationDataGV.Columns[3].Width = 380;
            QuotationDataGV.Columns[4].Width = 4;
            QuotationDataGV.Columns[5].Width = 80;
            QuotationDataGV.Columns[6].Width = 80;
            QuotationDataGV.Columns[7].Width = 80;

        }
        #endregion
                

        //生成报价单
        private void GenQuotationExcelBtn_Click(object sender, EventArgs e)
        {
            string excelFileName = ComprehensiveStaticClass.SaveExcleDialogShow();
            if (excelFileName == null)
            {
                return;
            }

            if (quotationExcelDT == null)
            {
                MessageBox.Show("无法生成！");
                return;
            }
            //生成QuotationExcel对应的DataTable
            //GenQuotationExcelDT(quotationExcelDT, quotationDGVDTStruct);
            //DataTable quotationExcelDT = new DataTable();
            GenQuotationExcelDT(quotationExcelDT, quotationDGVDTStruct);

            Match match = Regex.Match(evaluationReportClass.Line, @".*?\((.*?)\).*?",RegexOptions.Singleline);
            string layOut = match.Success ? match.Groups[1].Value : "";
            //quotationExcelDT导出为ExcelSheet
            ComprehensiveStaticClass.GenQuotationExcel(excelFileName,quotationDGVDTStruct, layOut);
        }


        /// <summary>
        /// 生成导出为quotationExcel的DataTable
        /// </summary>
        /// <param name="quotationExcelDT"></param>
        /// <param name="quotationDGVDTStruct"></param>
        private void GenQuotationExcelDT(DataTable quotationExcelDT, QuotationDGVDTStruct quotationDGVDTStruct)
        {
            #region //old is OK
            //quotationExcelDT的row
            DataRow dataRow;
            #region //添加头部信息
            dataRow = quotationExcelDT.NewRow();
            dataRow["DESCRIPTION"] = "FUJI SURFACE MOUNTING SYSTEM";
            quotationExcelDT.Rows.Add(dataRow);
            dataRow = quotationExcelDT.NewRow();
            dataRow["DESCRIPTION"] = "Fuji Scalable Placement Platform";
            quotationExcelDT.Rows.Add(dataRow);
            dataRow = quotationExcelDT.NewRow();
            dataRow["DESCRIPTION"] = "Model : NXT III";
            quotationExcelDT.Rows.Add(dataRow);
            #endregion

            #region //添加baseQotationDGVDT到quotationExcelDT
            for (int i = 0; i < quotationDGVDTStruct.baseQotationDGVDT.Rows.Count+1; i++)
            {
                if(i==0)//标题
                {
                    dataRow = quotationExcelDT.NewRow();
                    dataRow["DESCRIPTION"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["QuatationType"];
                    quotationExcelDT.Rows.Add(dataRow);
                    continue;
                }
                //quotationExcelDT.Rows.Add();
                //选中 且数量不为0
                if (Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["Option"]) == 0 && Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["QTY"]) != 0)
                {
                    dataRow = quotationExcelDT.NewRow();

                    dataRow["QTY"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["QTY"];
                    dataRow["unit"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["unit"];
                    dataRow["DESCRIPTION"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["DESCRIPTION"];
                    dataRow["UNIT PRICE"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["UNIT PRICE"];
                    dataRow["AMOUNT"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["AMOUNT"];
                    dataRow["EP"] = quotationDGVDTStruct.baseQotationDGVDT.Rows[i-1]["EP"];

                    quotationExcelDT.Rows.Add(dataRow);
                }
            }
            #endregion

            //int a = quotationExcelDT.Rows.Count;

            #region //添加 moduleHeadQotationDGVDT到DataTablr
            //模组head报价单部分 开始的位置

            for (int i = 0; i < quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows.Count; i++)
            {
                //if (i == 0)//首次开始位置
                //{
                //    //开始位置[quotationExcelDT.Rows.Count是下一行位置，+1+1，两行后开始] 结束位置[长度+标题一行]
                //    dTPoint_Struct.ModuleQotationDTStartEndRowList.Add(
                //        new int[2] {
                //            quotationExcelDT.Rows.Count+2,
                //            quotationExcelDT.Rows.Count+2+moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length+1
                //        });
                //    //添加插入扩过的两行
                //    quotationExcelDT.Rows.Add();
                //    quotationExcelDT.Rows[quotationExcelDT.Rows.Count - 1].ReadOnly = true;
                //    //quotationExcelDT[0,quotationExcelDT.Rows.Count - 1] = new DataGridViewButtonCell();

                //    quotationExcelDT.Rows.Add();
                //    quotationExcelDT.Rows[quotationExcelDT.Rows.Count - 1].ReadOnly = true;
                //    //quotationExcelDT[0, quotationExcelDT.Rows.Count - 1] = new DataGridViewButtonCell();
                //}
                //else
                //{
                //    dTPoint_Struct.ModuleQotationDTStartEndRowList.Add(
                //        new int[2] {
                //            quotationExcelDT.Rows.Count+1,
                //            quotationExcelDT.Rows.Count+1+moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length+1
                //        });
                //    //添加插入扩过的一行
                //    quotationExcelDT.Rows.Add();
                //    quotationExcelDT.Rows[quotationExcelDT.Rows.Count - 1].ReadOnly = true;
                //}
                //未选中或数量为0.不生成DataTable
                if (Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["Option"]) != 0 || Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"]) == 0)
                {
                    continue;
                }

                //跨过一行
                dataRow = quotationExcelDT.NewRow();
                quotationExcelDT.Rows.Add(dataRow);

                //插入DataGridView的位置
                for (int j = 0; j < quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length + 1; j++)
                {
                    dataRow = quotationExcelDT.NewRow();

                    if (j == 0)//首行，加标题
                    {
                        dataRow["DESCRIPTION"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QuatationType"];
                    }
                    else if (j == 1)
                    {
                        dataRow["QTY"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"];
                        dataRow["unit"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["unit"];
                        dataRow["DESCRIPTION"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n')[j - 1];
                        dataRow["UNIT PRICE"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["UNIT PRICE"];
                        dataRow["AMOUNT"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["AMOUNT"];
                        dataRow["EP"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["EP"];
                    }
                    else
                    {
                        dataRow["DESCRIPTION"] = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n')[j - 1];
                    }

                    quotationExcelDT.Rows.Add(dataRow);
                }

            }
            #endregion

            //跨过一行
            dataRow = quotationExcelDT.NewRow();
            quotationExcelDT.Rows.Add(dataRow);

            #region //添加 optionQotationDGVDT到DataGridView            
            //OptionDGVDT
            for (int i = 0; i < quotationDGVDTStruct.optionQotationDGVDT.Rows.Count+1; i++)
            {
                
                if (i==0)//标题行
                {
                    dataRow = quotationExcelDT.NewRow();
                    dataRow["DESCRIPTION"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i]["QuatationType"];
                    quotationExcelDT.Rows.Add(dataRow);
                    continue;
                }
                //选中且不为0
                if (Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["Option"]) == 0 && Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["QTY"]) != 0)
                {
                    dataRow = quotationExcelDT.NewRow();
                    dataRow["QTY"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["QTY"];
                    dataRow["unit"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["unit"];
                    dataRow["DESCRIPTION"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["DESCRIPTION"];
                    dataRow["UNIT PRICE"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["UNIT PRICE"];
                    dataRow["AMOUNT"] = quotationDGVDTStruct.optionQotationDGVDT.Rows[i-1]["AMOUNT"];
                    //dataRow["EP"].Value = OptionDGVDT.Rows[i-1]["EP"];
                    quotationExcelDT.Rows.Add(dataRow);
                }
            }
            #endregion

            //跨过一行
            dataRow = quotationExcelDT.NewRow();
            quotationExcelDT.Rows.Add(dataRow);

            #region //添加nozzleQotationDGVDTList 到DataTable
            foreach (DataTable nozzleQotationDGVDT in quotationDGVDTStruct.nozzleQotationDGVDTList)
            {
                //遍历，是否未全选中,默认未全选
                bool isChecked = false;
                for (int i = 0; i < nozzleQotationDGVDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(nozzleQotationDGVDT.Rows[i]["Option"]) == 0 && 
                        Convert.ToInt32(nozzleQotationDGVDT.Rows[i]["QTY"]) != 0)
                    {
                        isChecked = true;
                        break;
                    }
                }
                if (!isChecked)
                {//全部不写入DataTable(未选中)
                    continue;
                }
                for (int i = 0; i < nozzleQotationDGVDT.Rows.Count+1; i++)
                {
                    
                    //考虑全不选中是，仍会出现一个标题，特殊处理
                    if (i==0)//标题行
                {
                    dataRow = quotationExcelDT.NewRow();
                    dataRow["DESCRIPTION"] = nozzleQotationDGVDT.Rows[i]["QuatationType"];
                    quotationExcelDT.Rows.Add(dataRow);
                    continue;
                }
                    //选中且不为0
                    if (Convert.ToInt32(nozzleQotationDGVDT.Rows[i-1]["Option"]) == 0 && Convert.ToInt32(nozzleQotationDGVDT.Rows[i-1]["QTY"]) != 0)
                    {
                        dataRow = quotationExcelDT.NewRow();
                        dataRow["QTY"] = nozzleQotationDGVDT.Rows[i-1]["QTY"];
                        dataRow["unit"] = nozzleQotationDGVDT.Rows[i-1]["unit"];
                        dataRow["DESCRIPTION"] = nozzleQotationDGVDT.Rows[i-1]["DESCRIPTION"];
                        dataRow["UNIT PRICE"] = nozzleQotationDGVDT.Rows[i-1]["UNIT PRICE"];
                        dataRow["AMOUNT"] = nozzleQotationDGVDT.Rows[i-1]["AMOUNT"];
                        //dataRow["EP"].Value = OptionDGVDT.Rows[i-1]["EP"];
                        quotationExcelDT.Rows.Add(dataRow);
                    }
                }
            }
            #endregion

            //跨过一行
            dataRow = quotationExcelDT.NewRow();
            quotationExcelDT.Rows.Add(dataRow);

            #region //添加feederQotationDGVDTList 到DataTable
            foreach (DataTable feederQotationDGVDT in quotationDGVDTStruct.feederQotationDGVDTList)
            {
                //遍历，是否未全选中,默认未全选
                bool isChecked = false;
                for (int i = 0; i < feederQotationDGVDT.Rows.Count; i++)
                {
                    if (Convert.ToInt32(feederQotationDGVDT.Rows[i]["Option"]) == 0 &&
                        Convert.ToInt32(feederQotationDGVDT.Rows[i]["QTY"]) != 0)
                    {
                        isChecked = true;
                        break;
                    }
                }
                if (!isChecked)
                {//全部不写入DataTable(未选中)
                    continue;
                }
                for (int i = 0; i < feederQotationDGVDT.Rows.Count+1; i++)
                {
                    if (i==0)//标题行
                {
                    dataRow = quotationExcelDT.NewRow();
                    dataRow["DESCRIPTION"] = feederQotationDGVDT.Rows[i]["QuatationType"];
                    quotationExcelDT.Rows.Add(dataRow);
                    continue;
                }
                    //选中且不为0
                    if (Convert.ToInt32(feederQotationDGVDT.Rows[i-1]["Option"]) == 0 && Convert.ToInt32(feederQotationDGVDT.Rows[i-1]["QTY"]) != 0)
                    {
                        dataRow = quotationExcelDT.NewRow();
                        dataRow["QTY"] = feederQotationDGVDT.Rows[i-1]["QTY"];
                        dataRow["unit"] = feederQotationDGVDT.Rows[i-1]["unit"];
                        dataRow["DESCRIPTION"] = feederQotationDGVDT.Rows[i-1]["DESCRIPTION"];
                        dataRow["UNIT PRICE"] = feederQotationDGVDT.Rows[i-1]["UNIT PRICE"];
                        dataRow["AMOUNT"] = feederQotationDGVDT.Rows[i-1]["AMOUNT"];
                        //dataRow["EP"].Value = OptionDGVDT.Rows[i-1]["EP"];
                        quotationExcelDT.Rows.Add(dataRow);
                    }
                }
            }
            #endregion 
            #endregion

            #region //pass
            //DataRow dataRow;
            ////.Columns.Add("QTY");
            ////.Columns.Add("unit");
            // //.Columns.Add("DESCRIPTION");
            // //.Columns.Add();
            // //.Columns.Add("UNIT PRICE");
            // //.Columns.Add("AMOUNT");
            // //.Columns.Add("EP");
            //QuotationDataGV.Columns.Add("Select", "选择");
            ////添加其他列
            //QuotationDataGV.Columns.Add("QTY", "QTY");
            //QuotationDataGV.Columns.Add("unit", "");
            //QuotationDataGV.Columns.Add("DESCRIPTION", "DESCRIPTION");
            //QuotationDataGV.Columns.Add("", "");
            //QuotationDataGV.Columns.Add("UNIT PRICE", "UNIT PRICE");
            //QuotationDataGV.Columns.Add("AMOUNT", "AMOUNT");
            //QuotationDataGV.Columns.Add("EP", "EP");
            // //获取列数
            //int columnNum = QuotationDataGV.Columns.Count;
            //for (int i = 0; i < QuotationDataGV.Rows.Count; i++)
            //{
            //    dataRow = quotationExcelDT.NewRow();
            //    for (int j = 0; j < QuotationDataGV.Rows[i].Cells.Count; j++)
            //    {
            //        #region //base的行
            //        if (i >= dTPoint_Struct.baseQuotationDTStartEndRow[0] && i < dTPoint_Struct.baseQuotationDTStartEndRow[1])
            //        {
            //            //当前行是否选取，默认选取
            //            bool canRead = true;
            //            //获取当前行的选中状态
            //            if (j == 0 && QuotationDataGV.Rows[i].Cells[j] != null && QuotationDataGV.Rows[i].Cells[j].Value != null)
            //            {
            //                canRead = (bool)(((DataGridViewCheckBoxCell)QuotationDataGV.Rows[i].Cells[j]).Value);
            //                continue;
            //            }
            //            //非第一列
            //            if (j != 0)
            //            {
            //                //根据当前行是否读取 写入DataTable
            //                if (canRead)
            //                {
            //                    dataRow[j - 1] = QuotationDataGV.Rows[i].Cells[j].Value;
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        } 
            //        #endregion

            //        //Modules的行
            //        if (i >= dTPoint_Struct.ModuleQuotationDTStartEndRowList[0][0] && 
            //            i < dTPoint_Struct.ModuleQuotationDTStartEndRowList[dTPoint_Struct.ModuleQuotationDTStartEndRowList.Count-1][1])
            //        {
            //            //当前行是否选取，默认选取
            //            bool canRead = true;
            //            //获取当前modul的选中状态
            //            for (int i = 0; i < length; i++)
            //            {

            //            }
            //            if (j==0&& QuotationDataGV.Rows[i].Cells[j]!=null&& QuotationDataGV.Rows[i].Cells[j].Value!=null)
            //            {
            //                canRead = (bool)(((DataGridViewCheckBoxCell)QuotationDataGV.Rows[i].Cells[j]).Value);
            //                continue;
            //            }
            //            //非第一列
            //            if (j!=0)
            //            {
            //                //根据当前行是否读取 写入DataTable
            //                if (canRead)
            //                {
            //                    dataRow[j - 1] = QuotationDataGV.Rows[i].Cells[j].Value;
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //    }


            //当前行没有cell，即当前行空
            //标识，标识当前行是否为空行，默认是
            //bool isNullRow = true;
            //    for (int j = 0; j < QuotationDataGV.Rows[i].Cells.Count; j++)
            //    {
            //    //空 cell
            //    if (QuotationDataGV.Rows[i].Cells[j] == null || QuotationDataGV.Rows[i].Cells[j].Value != null)
            //    {
            //        continue;
            //    }
            //    else
            //    {
            //        //不是空行。
            //        isNullRow = false;
            //        if (j == 0)//选择第一列判断
            //        {
            //            //为空
            //            if (QuotationDataGV.Rows[i].Cells[j] == null || QuotationDataGV.Rows[i].Cells[j].Value != null)
            //            {

            //            }
            //        }
            //    }

            //        if (QuotationDataGV.Rows[i].Cells[j]!=null&& QuotationDataGV.Rows[i].Cells[j].Value!=null)
            //        {

            //        }
            //    }

            //} 
            #endregion
        }

        #region //DataGridView中显示DT
        /// <summary>
        /// 在QuotationDGVDT显示 quotationDGVDTStruct报价单
        /// </summary>
        /// <param name="quotationDGVDTStruct"></param>
        private void ShowQuotationDGVDT(QuotationDGVDTStruct quotationDGVDTStruct)
        {
            #region //quotation在DGV显示的实现。
            dTPoint_Struct = new DTPoint_Struct();
            dTPoint_Struct.baseQuotationDTStartEndRow = new int[2] {
                            0,
                            quotationDGVDTStruct.baseQotationDGVDT.Rows.Count-1
                        };
            #region //添加BaseQotationDataTable到DataGridView
            for (int i = 0; i < quotationDGVDTStruct.baseQotationDGVDT.Rows.Count; i++)
            {
                QuotationDataGV.Rows.Insert(dTPoint_Struct.baseQuotationDTStartEndRow[0] + i);
                if (Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["Option"]) == 0 && Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["QTY"]) != 0)
                {
                    QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells[0] = new DataGridViewCheckBoxCell();
                    ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells[0]).Value = true;
                }
                else if (Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["Option"]) != 0 || Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["QTY"]) == 0)
                {
                    QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells[0] = new DataGridViewCheckBoxCell();
                    ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells[0]).Value = false;
                }

                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["QTY"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["QTY"];
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["QTY"].Tag = Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["QTY"]);
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["unit"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["unit"];
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["DESCRIPTION"];
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["UNIT PRICE"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["UNIT PRICE"];
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["UNIT PRICE"].Tag = Convert.ToInt32(quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["UNIT PRICE"]);
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["AMOUNT"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["AMOUNT"];
                QuotationDataGV.Rows[dTPoint_Struct.baseQuotationDTStartEndRow[0] + i].Cells["EP"].Value = quotationDGVDTStruct.baseQotationDGVDT.Rows[i]["EP"];

            }
            #endregion

            //int a = QotationDataGV.Rows.Count;
            try
            {
                #region //添加 ModuleDGVDT到DataGridView
                //模组head报价单部分 开始的位置
                dTPoint_Struct.ModuleQuotationDTStartEndRowList = new List<int[]>();

                for (int i = 0; i < quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows.Count; i++)
                {
                    //if (i == 0)//首次开始位置
                    //{
                    //    //开始位置[QotationDataGV.Rows.Count是下一行位置，+1+1，两行后开始] 结束位置[长度+标题一行]
                    //    dTPoint_Struct.ModuleQotationDTStartEndRowList.Add(
                    //        new int[2] {
                    //            QotationDataGV.Rows.Count+2,
                    //            QotationDataGV.Rows.Count+2+moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length+1
                    //        });
                    //    //添加插入扩过的两行
                    //    QotationDataGV.Rows.Add();
                    //    QotationDataGV.Rows[QotationDataGV.Rows.Count - 1].ReadOnly = true;
                    //    //QotationDataGV[0,QotationDataGV.Rows.Count - 1] = new DataGridViewButtonCell();

                    //    QotationDataGV.Rows.Add();
                    //    QotationDataGV.Rows[QotationDataGV.Rows.Count - 1].ReadOnly = true;
                    //    //QotationDataGV[0, QotationDataGV.Rows.Count - 1] = new DataGridViewButtonCell();
                    //}
                    //else
                    //{
                    //    dTPoint_Struct.ModuleQotationDTStartEndRowList.Add(
                    //        new int[2] {
                    //            QotationDataGV.Rows.Count+1,
                    //            QotationDataGV.Rows.Count+1+moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length+1
                    //        });
                    //    //添加插入扩过的一行
                    //    QotationDataGV.Rows.Add();
                    //    QotationDataGV.Rows[QotationDataGV.Rows.Count - 1].ReadOnly = true;
                    //}
                    dTPoint_Struct.ModuleQuotationDTStartEndRowList.Add(
                            new int[2] {
                                        QuotationDataGV.Rows.Count+1,
                                        QuotationDataGV.Rows.Count+1+quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length+1
                            });
                    //添加插入扩过的一行
                    QuotationDataGV.Rows.Add();
                    QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;

                    //插入DataGridView的位置
                    for (int j = 0; j < quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n').Length + 1; j++)
                    {
                        QuotationDataGV.Rows.Insert(dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j);

                        if (j == 0 && Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["Option"]) == 0 && Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"]) != 0)
                        {//第一行 且 必选默认选中
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = true;
                        }
                        else if (j == 0 && (Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["Option"]) != 0 || Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"]) == 0))
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = false;
                            //((DataGridViewCheckBoxCell)QotationDataGV.Rows[dTPoint_Struct.ModuleQotationDTStartEndRowList[i][0] + j].Cells[0]).Value = false;
                        }
                        else//普通cell
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells[0].ReadOnly = true;
                        }
                        if (j == 0)//首行，加标题
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QuatationType"];
                        }
                        else if (j == 1)
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Tag = Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"]);
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["unit"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["unit"];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n')[j - 1];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["UNIT PRICE"];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Tag = Convert.ToInt32(quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["UNIT PRICE"]);
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["AMOUNT"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["AMOUNT"];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["EP"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["EP"];
                        }
                        else
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].ReadOnly = true;
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["DESCRIPTION"].ToString().Split('\n')[j - 1];
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].ReadOnly = true;
                            QuotationDataGV.Rows[dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] + j].Cells["AMOUNT"].ReadOnly = true;
                        }

                    }

                }
                #endregion

                #region //添加 OptionDGVDT到DataGridView
                dTPoint_Struct.OptionQuotationDTStartEndRow = new int[2] {
                            QuotationDataGV.Rows.Count+1,
                            QuotationDataGV.Rows.Count+1+quotationDGVDTStruct.optionQotationDGVDT.Rows.Count+1
                        };
                //添加插入扩过的1行
                QuotationDataGV.Rows.Add();
                QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;

                /////////OptionDGVDT
                ///否掉--由单一的拆分多为多个DataTable，将OptionDGVDT、feederDGVDT、nozzleDGVDT合在一块，使用QuatationType进行区分

                for (int i = 0; i < quotationDGVDTStruct.optionQotationDGVDT.Rows.Count + 1; i++)
                {
                    QuotationDataGV.Rows.Insert(dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i, 1);
                    if (i == 0)//第一行为标题
                    {
                        QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i]["QuatationType"];
                        QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].ReadOnly = true;
                        continue;
                    }
                    if (Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["Option"]) == 0 && Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["QTY"]) != 0)
                    {
                        QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells[0] = new DataGridViewCheckBoxCell();
                        ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells[0]).Value = true;
                    }
                    else if (Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["Option"]) != 0 || Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["QTY"]) == 0)
                    {
                        QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells[0] = new DataGridViewCheckBoxCell();
                        ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells[0]).Value = false;
                    }
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["QTY"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["QTY"];
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["QTY"].Tag = Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["QTY"]);
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["unit"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["unit"];
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["DESCRIPTION"];
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["UNIT PRICE"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["UNIT PRICE"];
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["UNIT PRICE"].Tag = Convert.ToInt32(quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["UNIT PRICE"]);
                    QuotationDataGV.Rows[dTPoint_Struct.OptionQuotationDTStartEndRow[0] + i].Cells["AMOUNT"].Value = quotationDGVDTStruct.optionQotationDGVDT.Rows[i - 1]["AMOUNT"];
                    //QotationDataGV.Rows[dTPoint_Struct.OptionQotationDTStartEndRow[0] + i].Cells["EP"].Value = OptionDGVDT.Rows[i-1]["EP"];

                }
                //int a = 0;
                #endregion

                #region //添加 NozzleDGVDT到DataGridView
                //nozzle报价单部分 开始的位置
                dTPoint_Struct.NozzleQuotationDTStartEndRowList = new List<int[]>();
                //添加nozzle报价单DataTableList
                for (int i = 0; i < quotationDGVDTStruct.nozzleQotationDGVDTList.Count; i++)
                {
                    dTPoint_Struct.NozzleQuotationDTStartEndRowList.Add(
                            new int[2] {
                                        QuotationDataGV.Rows.Count+1,
                                        QuotationDataGV.Rows.Count+1+quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows.Count+1//多个标题行
                            });
                    //添加插入跨过的一行
                    QuotationDataGV.Rows.Add();
                    QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;

                    for (int j = 0; j < quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows.Count + 1; j++)
                    {
                        if (j == 0)
                        {
                            //插入标题QuotationType
                            QuotationDataGV.Rows.Add();
                            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["DESCRIPTION"].Value =
                                quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[0]["QuatationType"];
                            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;
                            continue;
                        }
                        QuotationDataGV.Rows.Insert(dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j);
                        if (Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["Option"]) == 0
                            && Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["QTY"]) != 0)
                        {//默认 且数量不为0
                            QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = true;
                        }
                        else if (Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["Option"]) != 0 ||
                            Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["QTY"]) == 0)
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = false;
                        }
                        //填充数据
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Value = quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["QTY"];
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Tag = Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["QTY"]);
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["unit"].Value = quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["unit"];
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["DESCRIPTION"];
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Value = quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["UNIT PRICE"];
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Tag = Convert.ToInt32(quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["UNIT PRICE"]);
                        QuotationDataGV.Rows[dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] + j].Cells["AMOUNT"].Value = quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j - 1]["AMOUNT"];
                    }
                }
                #endregion
                #region //添加 FeederQuotationDGVDT到DataGridView
                //feeder报价单部分 开始的位置
                dTPoint_Struct.FeederQuotationDTStartEndRowList = new List<int[]>();
                //添加nozzle报价单DataTableList
                for (int i = 0; i < quotationDGVDTStruct.feederQotationDGVDTList.Count; i++)
                {
                    dTPoint_Struct.FeederQuotationDTStartEndRowList.Add(
                            new int[2] {
                                        QuotationDataGV.Rows.Count+1,
                                        QuotationDataGV.Rows.Count+1+quotationDGVDTStruct.feederQotationDGVDTList[i].Rows.Count+1//多个标题行
                            });
                    //添加插入跨过的一行
                    QuotationDataGV.Rows.Add();
                    QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;

                    for (int j = 0; j < quotationDGVDTStruct.feederQotationDGVDTList[i].Rows.Count + 1; j++)
                    {
                        if (j == 0)
                        {
                            //插入标题QuotationType
                            QuotationDataGV.Rows.Add();
                            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["DESCRIPTION"].Value =
                                quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[0]["QuatationType"];
                            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;
                            continue;
                        }
                        QuotationDataGV.Rows.Insert(dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j);
                        if (Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["Option"]) == 0
                            && Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["QTY"]) != 0)
                        {//默认 且数量不为0
                            QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = true;
                        }
                        else if (Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["Option"]) != 0 ||
                            Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["QTY"]) == 0)
                        {
                            QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells[0] = new DataGridViewCheckBoxCell();
                            ((DataGridViewCheckBoxCell)QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells[0]).Value = false;
                        }
                        //填充数据
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Value = quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["QTY"];
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["QTY"].Tag = Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["QTY"]);
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["unit"].Value = quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["unit"];
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["DESCRIPTION"].Value = quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["DESCRIPTION"];
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Value = quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["UNIT PRICE"];
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["UNIT PRICE"].Tag = Convert.ToInt32(quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["UNIT PRICE"]);
                        QuotationDataGV.Rows[dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] + j].Cells["AMOUNT"].Value = quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j - 1]["AMOUNT"];
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion

            #region//计算最低配置之和的实现和显示
            //添加最后一行
            QuotationDataGV.Rows.Add();
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["QTY"].Value = "Sub Total:";

            QuotationDataGV.Rows.Add();
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["QTY"].Value = "Export Packing:";

            QuotationDataGV.Rows.Add();
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].ReadOnly = true;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["QTY"].Value = "Total:FOB NAGOYA,JAPAN:";

            //求和
            minPrice = QuotationDataGVPriceSum(QuotationDataGV);
            #endregion

        }
        #endregion

        #region //求和的实现
        double QuotationDataGVPriceSum(DataGridView QuotationDataGV)
        {
            double SubTotal=0.0;
            double ExportPacking = 0.0;
            double TotalAll;
            for (int i = 0; i < QuotationDataGV.Rows.Count-3; i++)
            {
                if (QuotationDataGV.Rows[i]==null)
                {
                    continue;
                }
                //当前cell不为空，且勾选
                if (QuotationDataGV.Rows[i].Cells["AMOUNT"] != null && QuotationDataGV.Rows[i].Cells["AMOUNT"].Value!=null&&
                    !QuotationDataGV.Rows[i].Cells["AMOUNT"].Value.Equals(DBNull.Value))
                {
                    //是否选中
                    DataGridViewCheckBoxCell checkBoxCell = QuotationDataGV.Rows[i].Cells[0] as DataGridViewCheckBoxCell;
                    if (checkBoxCell != null&& (bool)checkBoxCell.Value==true)
                    {
                        SubTotal += Convert.ToDouble(QuotationDataGV.Rows[i].Cells["AMOUNT"].Value);
                    }
                }

                if (QuotationDataGV.Rows[i].Cells["EP"] != null && QuotationDataGV.Rows[i].Cells["EP"].Value!=null &&
                    !QuotationDataGV.Rows[i].Cells["EP"].Value.Equals(DBNull.Value))
                {
                    //是否选中
                    DataGridViewCheckBoxCell checkBoxCell = QuotationDataGV.Rows[i].Cells[0] as DataGridViewCheckBoxCell;
                    if (checkBoxCell != null && (bool)checkBoxCell.Value == true)
                    {
                        ExportPacking += Convert.ToDouble(QuotationDataGV.Rows[i].Cells["EP"].Value);
                    }
                }
                
            }

            TotalAll = SubTotal + ExportPacking;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 3].Cells["AMOUNT"].Value = SubTotal;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 2].Cells["AMOUNT"].Value = ExportPacking;
            QuotationDataGV.Rows[QuotationDataGV.Rows.Count - 1].Cells["AMOUNT"].Value = TotalAll;

            return TotalAll;
        }
        #endregion

        #region //折扣 discountComboBox數字變化事件
        private void discountComboB_TextChanged(object sender, EventArgs e)
        {
            //失去焦点后在继续
            //if (!discountComboB.Cursor)
            //{
            //    return;
            //}
            double discount;
            try
            {
                discount = Convert.ToDouble(discountComboB.Text) / (double)10;
                if (discount < 1)
                {
                    if (MessageBox.Show("折扣小于1，将修改原价的价格，是否确认？", "注意!", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        return;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("折扣必须为数字");
                discountComboB.Text = "10";
                return;
            }

            DisCountUnitPrice(discount);
        }
        #endregion

        #region //單價列折扣的實現
        /// <summary>
        /// 对UnitPrice打折
        /// </summary>
        /// <param name="discount"></param>
        private void DisCountUnitPrice(double discount)
        {
            //int priceIndx = 0;//价格对应的索引
            //for (int i = 0; i < QuotationDataGV.Rows.Count; i++)
            //{
            //    if (QuotationDataGV.Rows[i].Cells["UNIT PRICE"]!=null&&QuotationDataGV.Rows[i].Cells["UNIT PRICE"].Value != null)
            //    {
            //        QuotationDataGV.Rows[i].Cells["UNIT PRICE"].Value = unitPriceStandardList[priceIndx] * discount;
            //        priceIndx++;//每符合条件，赋值后+1
            //    }
            //}
            foreach (DataGridViewRow item in QuotationDataGV.Rows)
            {
                if (item.Cells["UNIT PRICE"]!=null&& item.Cells["UNIT PRICE"].Value!=null)
                {
                    item.Cells["UNIT PRICE"].Value = Convert.ToInt32(item.Cells["UNIT PRICE"].Tag) * discount;
                }                
            }
        }
        #endregion

        #region    //启用对feederNozzle数量的编辑
        private void feederNozzleQtyCheckB_CheckedChanged(object sender, EventArgs e)
        {
            if (feederNozzleQtyCheckB.Checked)
            {
                foreach (Control control in feederNozzlegroupB.Controls)
                {
                    control.Enabled = true;
                }
            }
            else
            {
                foreach (Control control in feederNozzlegroupB.Controls)
                {
                    CheckBox checkBox = control as CheckBox;
                    ComboBox comboBox = control as ComboBox;
                    if (comboBox != null)
                    {
                        comboBox.SelectedItem = comboBox.Items[0];
                    }
                    if (checkBox != null)
                    {
                        checkBox.Checked = false;
                    }
                    else
                    {
                        control.Enabled = false;
                    }
                }

                //调整的feederNozzle数量进行还原
            }
        }
        #endregion

        #region //cell 数量和单价修改事件方法
        private void QotationDataGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == null)
            {
                return;
            }
            //修改数值的有效性验证
            if ((e.ColumnIndex == 1 || e.ColumnIndex == 5)&&e.RowIndex<dataGridView.Rows.Count-3)
            {//Qty列和"UNIT PRICE"单价列，且当前行<行数-3
                dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].Value =
                    Convert.ToInt64(dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].Value) *
                    Convert.ToInt64(dataGridView.Rows[e.RowIndex].Cells["QTY"].Value);

                //当前DataGridView的值 复制到DataTable中
                //e.RowIndex的位置
                //先判定在属于哪个DT
                if (e.RowIndex >= dTPoint_Struct.baseQuotationDTStartEndRow[0] && e.RowIndex < dTPoint_Struct.baseQuotationDTStartEndRow[1])
                {
                    int num = 0;
                    //选中和取消 对应的索引
                    int indx = e.RowIndex - dTPoint_Struct.baseQuotationDTStartEndRow[0];
                    quotationDGVDTStruct.baseQotationDGVDT.Rows[indx]["QTY"] = int.TryParse(dataGridView.Rows[e.RowIndex].Cells["QTY"].Value.ToString(),out num)?num:throw new Exception("出现错误！");
                    quotationDGVDTStruct.baseQotationDGVDT.Rows[indx]["UNIT PRICE"] = int.TryParse(dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].Value.ToString(),out num)?num:throw new Exception("出现错误！");
                    quotationDGVDTStruct.baseQotationDGVDT.Rows[indx]["AMOUNT"] = int.TryParse(dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].Value.ToString(),out num)?num: throw new Exception("出现错误！");
                    //dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].ParseFormattedValue(dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].FormattedValue,)
                }
                else if (e.RowIndex >= dTPoint_Struct.ModuleQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.ModuleQuotationDTStartEndRowList[dTPoint_Struct.ModuleQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.ModuleQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前module报价单块，开始第一行即为选中的行。即当前第i个（moduleHeadQotationDGVDT的第i行）选中
                        if (dTPoint_Struct.ModuleQuotationDTStartEndRowList[i][0] == e.RowIndex)
                        {
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["QTY"] = dataGridView.Rows[e.RowIndex].Cells["QTY"].FormattedValue;
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["UNIT PRICE"] = dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].FormattedValue;
                            quotationDGVDTStruct.moduleHeadQotationDGVDT.Rows[i]["AMOUNT"] = dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].FormattedValue;
                        }
                    }
                }
                else if (e.RowIndex >= dTPoint_Struct.OptionQuotationDTStartEndRow[0] && e.RowIndex < dTPoint_Struct.OptionQuotationDTStartEndRow[1])
                {
                    //选中和取消 对应的索引-1,,有标题行
                    int indx = e.RowIndex - dTPoint_Struct.OptionQuotationDTStartEndRow[0] - 1;
                    quotationDGVDTStruct.optionQotationDGVDT.Rows[indx]["QTY"] = dataGridView.Rows[e.RowIndex].Cells["QTY"].FormattedValue;
                    quotationDGVDTStruct.optionQotationDGVDT.Rows[indx]["UNIT PRICE"] = dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].FormattedValue;
                    quotationDGVDTStruct.optionQotationDGVDT.Rows[indx]["AMOUNT"] = dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].FormattedValue;
                }
                else if (e.RowIndex >= dTPoint_Struct.NozzleQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.NozzleQuotationDTStartEndRowList[dTPoint_Struct.NozzleQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.NozzleQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前nozzle报价单块，选中的行。为当前QuatationType下的第 indx = e.RowIndex - dTPoint_Struct.FeederDTStartEndRow[0]-1索引
                        //多一行标题
                        //i 即为第i个QuatationType，即按QuatationType分组后的DataTableList的indx
                        if (e.RowIndex >= dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] &&
                            e.RowIndex < dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][1])
                        {
                            int indx = e.RowIndex - dTPoint_Struct.NozzleQuotationDTStartEndRowList[i][0] - 1;                            
                            quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[indx]["QTY"] = dataGridView.Rows[e.RowIndex].Cells["QTY"].FormattedValue;
                            quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[indx]["UNIT PRICE"] = dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].FormattedValue;
                            quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[indx]["AMOUNT"] = dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].FormattedValue;
                        }
                    }
                }
                else if (e.RowIndex >= dTPoint_Struct.FeederQuotationDTStartEndRowList[0][0] &&
                    e.RowIndex < dTPoint_Struct.FeederQuotationDTStartEndRowList[dTPoint_Struct.FeederQuotationDTStartEndRowList.Count - 1][1])
                {
                    for (int i = 0; i < dTPoint_Struct.FeederQuotationDTStartEndRowList.Count; i++)
                    {
                        //当前nozzle报价单块，选中的行。为当前QuatationType下的第 indx = e.RowIndex - dTPoint_Struct.FeederDTStartEndRow[0]-1索引
                        //多一行标题
                        //i 即为第i个QuatationType，即按QuatationType分组后的DataTableList的indx
                        if (e.RowIndex >= dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] &&
                            e.RowIndex < dTPoint_Struct.FeederQuotationDTStartEndRowList[i][1])
                        {
                            int indx = e.RowIndex - dTPoint_Struct.FeederQuotationDTStartEndRowList[i][0] - 1;
                            quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[indx]["QTY"] = dataGridView.Rows[e.RowIndex].Cells["QTY"].FormattedValue;
                            quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[indx]["UNIT PRICE"] = dataGridView.Rows[e.RowIndex].Cells["UNIT PRICE"].FormattedValue;
                            quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[indx]["AMOUNT"] = dataGridView.Rows[e.RowIndex].Cells["AMOUNT"].FormattedValue;
                        }
                    }
                }

                //重新计算sum
                QuotationDataGVPriceSum(QuotationDataGV);
            }
        }
        #endregion

        //Feeder倍数改变事件
        private void FeederQtyComboB_TextChanged(object sender, EventArgs e)
        {
            //if (FeederQtyComboB.Focused)
            //{
            //    return;
            //}
            NozzleQtyComboB.Text = FeederQtyComboB.Text;
            double mult;
            try
            {
                mult = Convert.ToDouble(FeederQtyComboB.Text);
                if (mult < 1)
                {
                    MessageBox.Show("倍数必须大于1");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("倍数必须为数字");
                FeederQtyComboB.Text = "1";
                return;
            }

            //

            MultQtys(dTPoint_Struct.FeederQuotationDTStartEndRowList[0][0],
                dTPoint_Struct.FeederQuotationDTStartEndRowList[dTPoint_Struct.FeederQuotationDTStartEndRowList.Count-1][1], mult);
            //for (int i = 0; i < FeederQtyIntsList.Count; i++)
            //{
            //    for (int j = 0; j < FeederQtyIntsList[i].Length; j++)
            //    {
            //        quotationDGVDTStruct.feederQotationDGVDTList[i].Rows[j]["QTY"] = Math.Ceiling(FeederQtyIntsList[i][j] * mult);
            //    }
            //}
        }

        //Nozzle 倍数改变事件
        private void NozzleQtyComboB_TextChanged(object sender, EventArgs e)
        {
            //if (NozzleQtyComboB.Focused)
            //{
            //    return;
            //}
            FeederQtyComboB.Text = NozzleQtyComboB.Text;
            double mult;
            try
            {
                mult = Convert.ToDouble(NozzleQtyComboB.Text);
                if (mult < 1)
                {
                    MessageBox.Show("倍数必须大于1");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("倍数必须为数字");
                NozzleQtyComboB.Text = "1";
                return;
            }

            MultQtys(dTPoint_Struct.NozzleQuotationDTStartEndRowList[0][0], 
                dTPoint_Struct.NozzleQuotationDTStartEndRowList[dTPoint_Struct.NozzleQuotationDTStartEndRowList.Count-1][1], mult);

            //for (int i = 0; i < NozzleQtyIntsList.Count; i++)
            //{
            //    for (int j = 0; j < NozzleQtyIntsList[i].Length; j++)
            //    {
            //        quotationDGVDTStruct.nozzleQotationDGVDTList[i].Rows[j]["QTY"] =Math.Ceiling(NozzleQtyIntsList[i][j] * mult);
            //    }
            //}
        }

        ////修改指定范围的Qty数量，即feeder和Nozzle的倍数
        void MultQtys(int DGVStartRow,int DGVEndRow,double mult)
        {
            for (int i = DGVStartRow; i < DGVEndRow; i++)
            {
                if (QuotationDataGV.Rows[i].Cells["QTY"] != null &&
                    QuotationDataGV.Rows[i].Cells["QTY"].Value != null )
                {
                    QuotationDataGV.Rows[i].Cells["QTY"].Value = Math.Ceiling(Convert.ToInt32(QuotationDataGV.Rows[i].Cells["QTY"].Tag) * mult);
                }
            }
            

            //for (int i = 0; i < baseOriginQty.Count; i++)
            //{
            //    if (QuotationDataGV.Rows[DGVStartRow + i].Cells["QTY"] == null &&
            //        QuotationDataGV.Rows[DGVStartRow + i].Cells["QTY"].Value == null&& baseOriginQty[i]!=-1)
            //    {
            //        MessageBox.Show("处理出现了错误，请停止使用报价单模块！");
            //        return;
            //    }
            //        if (QuotationDataGV.Rows[DGVStartRow+i].Cells["QTY"] == null&&
            //        QuotationDataGV.Rows[DGVStartRow+i].Cells["QTY"].Value == null)
            //    {
            //        continue;
            //    }
            //    QuotationDataGV.Rows[DGVStartRow + i].Cells["QTY"].Value = baseOriginQty[i] * mult;
            //}
        }
        //直接修改DGVDT的数值

        //private void QotationDataGV_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        //{
        //    if (e.ColumnIndex !=0|| e.RowIndex<0)
        //    {
        //        return;
        //    }
        //    DataGridView dataGridView = sender as DataGridView;
        //    if (e.RowIndex>= dataGridView.Rows.Count)
        //    {
        //        return;
        //    }
        //    if (((DataGridViewCheckBoxCell)dataGridView[e.ColumnIndex, e.RowIndex]).ReadOnly)
        //    {
        //        //((DataGridViewCheckBoxCell)dataGridView[e.ColumnIndex, e.RowIndex]).
        //        //ControlPaint.DrawBorder(e.Graphics, e.CellBounds, SystemColors.Control, ButtonBorderStyle.None);

        //        ControlPaint.DrawButton(e.Graphics, e.CellBounds, ButtonState.Flat);
        //    }
        //}
    }
}
