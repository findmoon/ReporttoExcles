using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Threading;
using System.Data.SQLite;

namespace Exicel转换1
{
    public partial class XmlReport : UserControl
    {
        public XmlReport()
        {
            InitializeComponent();
        }
        ////窗体的显示类型转换，
        ////报错未实现该方法
        //public static explicit operator XmlReport(Form v)
        //{
        //    throw new NotImplementedException();
        //}

        int openFolderNum = 0;

        //模组代码对应的模组类型的 字典
        Dictionary<string, string> moduleCodeDict = new Dictionary<string, string>()
        {
            {"601","M3III" },
            {"602","M6III" }
        };
        //Head代码对应的head类型的 字典
        Dictionary<string, string> headCodeDict = new Dictionary<string, string>()
        {
            {"15","H24" },
            {"18","H04SF" }
        };

        //存放生产模式的变量
        string productionMode = null;
        public string ProductionMode
        {
            set
            {
                productionMode = value;
            }
        }

        //最后包含summary和layout(layouInfoDT feederNozzleDT)的信息的list，实现可以打开多个export文件夹
        List<BaseComprehensive> summaryandLayout_ComprehensiveList = new List<BaseComprehensive>();

        //最后包含summaryInfo和layouInfoDT feederNozzleDT的信息
        BaseComprehensive theEndSummaryandLayout = null;


        //已打开开的目录
        string pathXmlFolder = null;

        #region //打开xml report文件夹 按钮
        private void OpenXmlFolder_Btn_Click(object sender, EventArgs e)
        {
            #region //将每次打开一个文件夹都会变化的变量设置为局部变量，全局的话会影响多次的打开后的判断
            //Feeder type : num
            XmlDocument xml_FeederReportUnit = null;

            //module：Header type
            //public XmlDocument xml_HeadReportUnit = new XmlDocument();

            //Nozzle (ncstNzlName): count
            XmlDocument xml_NozzleChangerReportUnit = null;

            //module : head cycletime，置件数量Qty
            //只有二轨的time
            XmlDocument xml_TimingReportUnitNxt = null;

            //machine name\job name\top or bottom
            XmlDocument xml_TimingReportHeader = null;

            //seqBrdNum=>borad num，
            XmlDocument xml_PartReportUnit = null;

            //pannel size   prgPnlLength
            XmlDocument xml_PartReportHead = null;

            //xml_TimingReportHeaderUnit   Total_Cycle_Time  Number_of_Placements
            XmlDocument xml_TimingReportHeaderUnit = null;

            //rtf文件的StreamReader，用于读取里面内容
            StreamReader srRTFFile = null;
            #endregion

            //打开文件夹浏览器
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                pathXmlFolder = folderBrowserDialog1.SelectedPath;

                /*获取多个类型格式的文件
                 * System.IO.Directory.GetFiles("c:\","(*.jpg|*.bmp)"); 
                 * var files = Directory.GetFiles("C:\\path", "*.*", SearchOption.AllDirectories)
                    .Where(s => s.EndsWith(".bmp") || s.EndsWith(".jpg"));
                 */

                var files = Directory.GetFiles(pathXmlFolder, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(s => s.EndsWith(".xml", true, null) || s.EndsWith(".rtf", true, null)).ToArray();

                #region 使用构造DirectoryInfo类，获取所有信息
                ////获取路径下所有的 xml和rtf文件
                //FileInfo[] xmlrtfFileinfos = new DirectoryInfo(pathXmlFolder).GetFiles("*.xml", SearchOption.TopDirectoryOnly);
                ////获各个xml文件，并转换为XmlDocuumentduixiang
                ////获取每个路径并赋值给xml                
                //foreach (FileInfo xmlrtfFileinfo in xmlrtfFileinfos)
                //{
                //    //加载xml
                //    if (xmlrtfFileinfo.FullName.IndexOf(".xml")>0)
                //    {
                //        XmlDocument xmlDocument = new XmlDocument();
                //        xmlDocument.Load(xmlrtfFileinfo.FullName);

                //        switch (xmlDocument.DocumentElement.Name)//获取根节点DocumentElement
                //        {
                //            case "TimingReportHeaderUnit":
                //                xml_TimingReportHeaderUnit = xmlDocument; break;
                //            case "TimingReportUnit":
                //                xml_TimingReportUnitNxt = xmlDocument; break;
                //            case "TimingReportHeader":
                //                xml_TimingReportHeader = xmlDocument; break;
                //            case "PartReportUnit":
                //                xml_PartReportUnit = xmlDocument; break;
                //            case "PartReportHead":
                //                xml_PartReportHead = xmlDocument; break;
                //            case "NozzleChangerReportUnit":
                //                xml_NozzleChangerReportUnit = xmlDocument; break;
                //            case "FeederReportUnit":
                //                xml_FeederReportUnit = xmlDocument; break;
                //            default:                                
                //                break;
                //        }
                //        //不能使用包含判断，名字中本身就有包含
                //        //if (xmlFileinfo.FullName.Contains("FeederReportUnit"))
                //        //{
                //        //    xml_FeederReportUnit.Load(xmlFileinfo.FullName);
                //        //}
                //        //if (xmlFileinfo.FullName.Contains("HeadReportUnit"))
                //        //{
                //        //    xml_HeadReportUnit.Load(xmlFileinfo.FullName);
                //        //}
                //        //if (xmlFileinfo.FullName.Contains("NozzleChangerReportUnit"))
                //        //{
                //        //    xml_NozzleChangerReportUnit.Load(xmlFileinfo.FullName);
                //        //}

                //        //if (xmlFileinfo.FullName.Contains("TimingReportUnitReport"))
                //        //{
                //        //    xml_TimingReportUnitNxt.Load(xmlFileinfo.FullName);
                //        //}

                //        //if (xmlFileinfo.FullName.Contains("TimingReportHeader"))
                //        //{
                //        //    xml_TimingReportHeader.Load(xmlFileinfo.FullName);
                //        //}

                //        //if (xmlFileinfo.FullName.Contains("PartReportUnit"))
                //        //{
                //        //    xml_PartReportUnit.Load(xmlFileinfo.FullName);
                //        //}

                //        //if (xmlFileinfo.FullName.Contains("PartReportHead"))
                //        //{
                //        //    xml_PartReportHead.Load(xmlFileinfo.FullName);
                //        //}

                //        //如何判断XMLDocument是否为空 ||或xmlDocument文档个数不够时不执行，重新选择                   
                //    }
                //    else//加载其他文件
                //    {
                //        //存在多个rtf时终止运行

                //        //创建文件流
                //        FileStream fs = new FileStream(xmlrtfFileinfo.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
                //        //读取流
                //        srRTFFile = new StreamReader(fs);
                //    }
                //}

                #endregion

                int rtfNum = 0;//记录rtf文件个数
                for (int i = 0; i < files.Length; i++)
                {
                    if (files[i].IndexOf(".xml") > 0)
                    {
                        XmlDocument xmlDocument = new XmlDocument();
                        xmlDocument.Load(files[i]);

                        switch (xmlDocument.DocumentElement.Name)//获取根节点DocumentElement
                        {
                            case "TimingReportHeaderUnit":
                                xml_TimingReportHeaderUnit = xmlDocument; break;
                            case "TimingReportUnit":
                                xml_TimingReportUnitNxt = xmlDocument; break;
                            case "TimingReportHeader":
                                xml_TimingReportHeader = xmlDocument; break;
                            case "PartReportUnit":
                                xml_PartReportUnit = xmlDocument; break;
                            case "PartReportHead":
                                xml_PartReportHead = xmlDocument; break;
                            case "NozzleChangerReportUnit":
                                xml_NozzleChangerReportUnit = xmlDocument; break;
                            case "FeederReportUnit":
                                xml_FeederReportUnit = xmlDocument; break;
                            default:
                                break;
                        }
                    }
                    else//加载其他文件
                    {
                        //存在多个rtf时终止运行
                        if (++rtfNum > 1)
                        {
                            MessageBox.Show("存在多个rtf文件，请确认与当前报告相对应的rtf文件！\n然后重新选择！");
                            if (srRTFFile != null)
                            {
                                srRTFFile.Dispose();
                            }
                            if (xml_TimingReportHeaderUnit != null)
                            {
                                //xml已经存在时的销毁
                            }
                            return;
                        }

                        //创建文件流
                        FileStream fs = new FileStream(files[i], FileMode.Open, FileAccess.Read, FileShare.Read);
                        //读取流
                        srRTFFile = new StreamReader(fs);
                    }
                }
                //打开的目录和目录数
                openFolderNum++;
                HasOpenFolderLbel.Text += pathXmlFolder + "\n";

            }
            else
            {
                return;//未选择任何
            }
            //生成最后的BaseComprehensive，未实现生成BaseComprehensive List
            GenBaseComprehensive(xml_TimingReportHeaderUnit,xml_TimingReportUnitNxt,xml_TimingReportHeader,
            xml_PartReportUnit,xml_PartReportHead,xml_NozzleChangerReportUnit,xml_FeederReportUnit,srRTFFile);
        }
        #endregion


        #region //生成 summaryandLayout_ComprehensiveList  theEndSummaryandLayout
        //生成所有信息 BaseComprehensive class 由summaryinfo和layoutinfo组成
        public void GenBaseComprehensive(XmlDocument xml_TimingReportHeaderUnit,XmlDocument xml_TimingReportUnitNxt,
            XmlDocument xml_TimingReportHeader,XmlDocument xml_PartReportUnit,XmlDocument xml_PartReportHead,
            XmlDocument xml_NozzleChangerReportUnit,XmlDocument xml_FeederReportUnit,StreamReader srRTFFile)
        {
            //生成综合的类，包含summary和layout及需要的验证信息。
            BaseComprehensive baseComprehensive = new BaseComprehensive();

            string jobName = xml_TimingReportHeader.GetElementsByTagName("JobName")[1].InnerText;
            string t_or_b = xml_TimingReportHeader.GetElementsByTagName("BoardSide")[1].InnerText;
            string machine_name = xml_TimingReportHeader.GetElementsByTagName("McName")[1].InnerText;

            int MachineCount = xml_TimingReportUnitNxt.GetElementsByTagName("Module").Count - 1;

            XmlNodeList boardNodeList = xml_PartReportUnit.GetElementsByTagName("seqBrdNum");
            int boardQty = Int32.Parse(boardNodeList[1].InnerText);
            for (int i = 1; i < boardNodeList.Count; i++)
            {
                if (boardQty < Int32.Parse(boardNodeList[i].InnerText))
                {
                    boardQty = Int32.Parse(boardNodeList[i].InnerText);
                }

            }

            //使用元组存放module head cycletime  Qt
            //获取 line2_module_head_cycletime_qty_TupleList 元组列表，使用sqlite
            //调用方法
            //module、head、avgCPH、prior_productionCPH	、high_precisionCPH做成一个小DataTable输出.
            //使用out参数
            DataTable module_Head_TheoryCPH_Table = null;
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList =
                Getline2_module_head_cycletime_qty_TupleList(xml_TimingReportUnitNxt,out module_Head_TheoryCPH_Table);

            if (module_Head_TheoryCPH_Table == null)
            {
                MessageBox.Show("获取CPH出现了严重错误，请联系作者！");
                return ;
            }


            //模组总数

            //模组统计
            //字典，存放<模组类型:模组个数>
            Dictionary<string, int> module_Statistics;
            //创建统计模组信息的字典
            module_Statistics = new Dictionary<string, int>();
            //初始化，仅记录M3III，M6III，初始为0

            #region 统计模组信息
            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                if (module_Statistics.ContainsKey(line2_module_head_cycletime_qty_TupleList[j].Item1))
                {
                    module_Statistics[line2_module_head_cycletime_qty_TupleList[j].Item1]++;
                }
                else
                {
                    module_Statistics.Add(line2_module_head_cycletime_qty_TupleList[j].Item1, 1);
                }
            }
            #endregion

            #region machineKid
            //依据M3III|M6III，判断是NXTIII还是其他，生成machinekind
            string machineKind = "";
            if (module_Statistics.Keys.Contains<string>("M3III") || module_Statistics.Keys.Contains<string>("M6III"))
            {
                machineKind += "NXTIII";
            }
            if (module_Statistics.Keys.Contains<string>("M3II") || module_Statistics.Keys.Contains<string>("M6II"))
            {
                machineKind += "NXTII";
            }
            if (module_Statistics.Keys.Contains<string>("M3S") || module_Statistics.Keys.Contains<string>("M6S") ||
                module_Statistics.Keys.Contains<string>("M3") || module_Statistics.Keys.Contains<string>("M6"))
            {
                machineKind += "NXT";
            }
            #endregion

            #region//拼接Line字符串

            //重新计算base信息
            //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
            //base尽可能少的分配
            //模组 M的数量
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
            string Line = "";
            string Line_short = "";
            foreach (var item in module_Statistics)
            {
                Line_short += item.Key + "*" + item.Value.ToString() + "+";
            }
            Line_short = machineKind + "-(" + Line_short.Substring(0, Line_short.Length - 1) + ")";
            //判断4Mhe 2M base的数量，进行拼接
            if (baseCount_2M == 0 && baseCount_4M != 0)
            {
                Line = Line_short + "\n" + "4MBASE III * " + baseCount_4M.ToString();
            }
            if (baseCount_2M != 0 && baseCount_4M != 0)
            {
                Line = Line_short + "\n" + "4MBASE III * " + baseCount_4M.ToString() + "+2MBASE III * " + baseCount_2M.ToString();
            }
            if (baseCount_2M != 0 && baseCount_4M == 0)
            {
                Line = Line_short + "\n" + "2MBASE III * " + baseCount_2M.ToString();
            }
            #endregion

            #region//获取组合"Head Type"的信息

            //统计头的信息 字典
            Dictionary<string, int> Head_Statistics;
            Head_Statistics = new Dictionary<string, int>();

            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                // 只有两种情况，包含key和不包含key
                if (Head_Statistics.ContainsKey(line2_module_head_cycletime_qty_TupleList[j].Item2))
                {
                    Head_Statistics[line2_module_head_cycletime_qty_TupleList[j].Item2] += 1;
                }
                else //if (!Head_Statistics.ContainsKey(allHeadTypeString[j]))
                {
                    Head_Statistics.Add(line2_module_head_cycletime_qty_TupleList[j].Item2, 1);
                }

            }

            //拼接Head Type
            string Head_Type = "";
            foreach (var item in Head_Statistics)
            {
                Head_Type += item.Key + "*" + item.Value + "+";
            }
            Head_Type = Head_Type.Substring(0, Head_Type.Length - 1);
            // MessageBox.Show(Head_Type);
            #endregion

            #region Pannelsize    NXT4M_PartReportHead_T----> prgPnlLength  prgPnlWidth
            //pannelsize取整
            string pannelLength = string.Format("{0:0}",
                Convert.ToDouble(xml_PartReportHead.GetElementsByTagName("prgPnlLength")[1].InnerText));
            string pannelWidth = string.Format("{0:0}",
                Convert.ToDouble(xml_PartReportHead.GetElementsByTagName("prgPnlWidth")[1].InnerText));
            string Pannelsize = pannelLength + "*" + pannelWidth;
            #endregion

            #region placementNumber    NXT4M_TimingReportHeaderUnit_T----> Number_of_Placements
            string placementNumber = xml_TimingReportHeaderUnit.GetElementsByTagName("Number_of_Placements")[1].InnerText;
            #endregion

            #region Remark  电气功率ip等。//未加入try LTCTry
            #region //生成Remake信息
            //
            double gonglv = Count_M3 * ComprehensiveStaticClass.power_Dictionary["M3III"] +
                Count_M6 * ComprehensiveStaticClass.power_Dictionary["M6III"] +
                baseCount_2M * ComprehensiveStaticClass.power_Dictionary["2MBase"] +
                baseCount_4M * ComprehensiveStaticClass.power_Dictionary["4MBase"];
            int haoqiliang = baseCount_4M * ComprehensiveStaticClass.air_Consumption_Dictionary["4MBase"] +
                baseCount_2M * ComprehensiveStaticClass.air_Consumption_Dictionary["2MBase"];
            int countu_ip = baseCount_2M + baseCount_4M;
            double baseLength = ComprehensiveStaticClass.baseLength_Dictionary["2MBase"] * baseCount_2M +
                ComprehensiveStaticClass.baseLength_Dictionary["4MBase"] * baseCount_4M;
            //最后拼接
            string remark = string.Format("功率：{0}\n耗气量：{1}\n长度：{2}\n最少IP数：{3}",
                gonglv, haoqiliang, baseLength, countu_ip);
            #endregion
            #endregion

            

            #region //生成ExpressSummayDataTable
            //cPH 两个轨道综合的CPH，cycletime
            DataTable expressSummayDataTable = ComprehensiveStaticClass.getExpressSummayDataTable(boardQty, Line,
                Head_Type, jobName, Pannelsize, placementNumber, remark);
            //baseComprehensive.GetTheEndInfoDT()
            //SummaryEveryInfo.PictureDataByte = genJobsPicture(allModuleTypeString, module_Statistics);
            #endregion

            #region //获取Feeder type、feederQty
            Dictionary<string, int> feeder_Statistics = new Dictionary<string, int>();
            XmlNodeList feeder_xmlNodeList = xml_FeederReportUnit.GetElementsByTagName("Unit");
            for (int i = 0; i < feeder_xmlNodeList.Count; i++)
            {
                XmlNode feederName = feeder_xmlNodeList[i].SelectSingleNode("fsFdrName");
                if (feeder_Statistics.ContainsKey(feederName.InnerText))
                {
                    feeder_Statistics[feederName.InnerText]++;
                }
                else
                {
                    feeder_Statistics.Add(feederName.InnerText, 1);
                }
            }
            #endregion

            #region //获取Nozzle type、Nozzle Qty
            Dictionary<string, int> nozzle_Statistics = new Dictionary<string, int>();
            XmlNodeList nozzle_xmlNodeList = xml_NozzleChangerReportUnit.GetElementsByTagName("Unit");
            for (int i = 0; i < nozzle_xmlNodeList.Count; i++)
            {
                XmlNode nozzleName = nozzle_xmlNodeList[i].SelectSingleNode("./ncstNzlName");
                if (nozzle_Statistics.ContainsKey(nozzleName.InnerText))
                {
                    nozzle_Statistics[nozzleName.InnerText]++;
                }
                else
                {
                    nozzle_Statistics.Add(nozzleName.InnerText, 1);
                }
            }
            #endregion

            //获取生产模式
            //Thread getProductionModelThread = new Thread(GetProductionModel);
            //getProductionModelThread.IsBackground = true;
            //getProductionModelThread.Start();
            GetProductionModel();

            //获取 genLayoutEndDT，最后的cycletime cph theory_cph layoutDT、feedernozzleDT
            Tuple<double, int, DataTable, DataTable> layoutEndInfo_tuple = genLayoutEndDT(srRTFFile,productionMode,
                line2_module_head_cycletime_qty_TupleList, feeder_Statistics, nozzle_Statistics);
            //判断是否获取到 
            if (layoutEndInfo_tuple == null)
            {
                return;
            }

            /// Output Board/hour  -----> board qty *3600/cycletime
            // Output Board/Day(22H)  ---->(Output Board/hour)*22
            //CPH Rate   ----> CPH/理论CPH
            // CPP  链接报价单，未实现
            //以上由函数GetTheEndSummaryInfoDT实现

            //所有模组type的string。获取最后的图片 byte
            string[] allModuleTypeString = new string[line2_module_head_cycletime_qty_TupleList.Count];
            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                allModuleTypeString[j] = line2_module_head_cycletime_qty_TupleList[j].Item1;
            }
            #region //pictureData放在前面，未生成则不进行下面对象的创建
            //判断是否获取成功图片的byte[]
            byte[] pictureDataByte = ComprehensiveStaticClass.genJobsPicture(allModuleTypeString, module_Statistics);
            if (pictureDataByte == null)
            {
                return;
            }
            #endregion

            //获取最后的综合的 SummaryandLayout
            theEndSummaryandLayout = new BaseComprehensive();
            theEndSummaryandLayout.T_or_B = t_or_b;
            theEndSummaryandLayout.MachineCount = MachineCount;
            theEndSummaryandLayout.MachineKind = machineKind;
            theEndSummaryandLayout.Jobname = jobName;
            //获取CPHRate
            //dr["module"] 
            //dr["Head Type"] 
            //dr["avgCPH"]
            //dr["prior_productionCPH"]
            //dr["high_precisionCPH"]
            //dr["special"]
            int theoryCPH = 0;
            for (int i = 0; i < module_Head_TheoryCPH_Table.Rows.Count; i++)
            {
                //判断高速生产模式cph是否为 空
                if (!module_Head_TheoryCPH_Table.Rows[i]["prior_productionCPH"].Equals(0))
                {
                    theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["prior_productionCPH"]);
                }
                else if (!module_Head_TheoryCPH_Table.Rows[i]["avgCPH"].Equals(0))
                {
                    theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["avgCPH"]);
                }
                else if (!module_Head_TheoryCPH_Table.Rows[i]["high_precisionCPH"].Equals(0))
                {
                    theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["high_precisionCPH"]);
                }
                
            }
            string cphRate = string.Format("{0:P}", System.Convert.ToDouble(layoutEndInfo_tuple.Item2) / System.Convert.ToDouble(theoryCPH));
            // machine_name
            //获取最后的SummaryinfoDT
            theEndSummaryandLayout.GetTheEndSummaryInfoDT(expressSummayDataTable,
                layoutEndInfo_tuple.Item1, layoutEndInfo_tuple.Item2,cphRate);
                        
            theEndSummaryandLayout.PictureDataByte = pictureDataByte;

            theEndSummaryandLayout.layoutDT = layoutEndInfo_tuple.Item3;
            theEndSummaryandLayout.feederNozzleDT = layoutEndInfo_tuple.Item4;
            theEndSummaryandLayout.AllModuleType = allModuleTypeString;

            //所有头Head type的string
            string[] allHeadTypeString = new string[line2_module_head_cycletime_qty_TupleList.Count];
            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                allHeadTypeString[j] = line2_module_head_cycletime_qty_TupleList[j].Item2;
            }
            theEndSummaryandLayout.AllHeadType = allHeadTypeString;

            //添加到SummaryandLayout_ComprehensiveList列表
            summaryandLayout_ComprehensiveList.Add(theEndSummaryandLayout);
            //使用全局变量，因为多处需要判断全局变量的值，因此如果选择一个文件夹后处理完所有的过程后，需要重置所有变量，
            //否则导致下一次选择判断的混乱。
        }

        #endregion

        public List<Tuple<string, string, string, string>> Getline2_module_head_cycletime_qty_TupleList(XmlDocument xml_TimingReportUnitNxt,
            out DataTable module_Head_TheoryCPH_Table)
        {
            //使用元组存放module head cycletime  Qty
            //获取 line2_module_head_cycletime_qty_TupleList 元组列表，使用sqlite
            //调用方法
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList =
                new List<Tuple<string, string, string, string>>();

            //获取line对应的module\head\Qty|cycletime信息 
            XmlNodeList truUnitNode = xml_TimingReportUnitNxt.GetElementsByTagName("Unit");
            

            //初始化module_Head_TheoryCPH_Table对象
            module_Head_TheoryCPH_Table = new DataTable();
            //DataColumn dc = new DataColumn();
            //dc.
            module_Head_TheoryCPH_Table.Columns.Add("module");
            module_Head_TheoryCPH_Table.Columns.Add("Head Type");
            module_Head_TheoryCPH_Table.Columns.Add("avgCPH"); 
            module_Head_TheoryCPH_Table.Columns.Add("prior_productionCPH");
            module_Head_TheoryCPH_Table.Columns.Add("high_precisionCPH");
            module_Head_TheoryCPH_Table.Columns.Add("special");

            //定义dqlite的数据库连接对象

            string connString = string.Format("Data Source={0};Version=3;", @".\convertDB.db");
            SQLiteConnection liteConn = new SQLiteConnection();//创建数据库连接实例
            liteConn.ConnectionString = connString;
            try
            {
                liteConn.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("连接失败");
            }
           
            for (int i = 1; i < truUnitNode.Count; i++)//从1开始读取
            {                
                try
                {
                    //
                    string moduleid = truUnitNode[i].SelectSingleNode("./cfgModuleType1").InnerText;
                    string headid = truUnitNode[i].SelectSingleNode("./cfgHeadType1_1").InnerText;
                    
                    //sqlite的命令对象,查询moduleType
                    SQLiteCommand command = new SQLiteCommand(string.Format("select moduletype from ModuleTable where moduleid={0};", moduleid),
                        liteConn);
                    if (liteConn.State == ConnectionState.Closed)
                    {
                        liteConn.Open();
                    }
                    string moduleType = command.ExecuteScalar().ToString();
                    

                    //查询headType
                    command = new SQLiteCommand(string.Format("select head from HeadTable where headid={0};", headid),
                        liteConn);
                    if (liteConn.State == ConnectionState.Closed)
                    {
                        liteConn.Open();
                    }
                    string head = command.ExecuteScalar().ToString();
                    

                    //查询CPH
                    string sqlQueryCPH = string.Format(
                        "select a.moduletype,b.head,c.avgCPH,c.prior_productionCPH,c.high_precisionCPH,c.special from " +
                        "ModuleTable a,HeadTable b ,CPHTable c where a.moduleid=c.moduleid and b.headid=c.headid and" +
                        " c.headid={0} and c.moduleid={1};",
                        headid, moduleid);
                    command = new SQLiteCommand(sqlQueryCPH, liteConn);
                    DataSet ds = new DataSet();
                    if (liteConn.State == ConnectionState.Closed)
                    {
                        liteConn.Open();
                    }
                    SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                    da.Fill(ds);

                    //datarow module_Head_TheoryCPH_Table
                    DataRow dr = module_Head_TheoryCPH_Table.NewRow();
                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                    {
                        if (j>0)//先跳过H24 S G head 的处理
                        {
                            continue;
                        }
                        dr["module"] = ds.Tables[0].Rows[j]["moduletype"];
                        dr["Head Type"] = ds.Tables[0].Rows[j]["head"];
                        dr["avgCPH"] = ds.Tables[0].Rows[j]["avgCPH"];
                        dr["prior_productionCPH"] = ds.Tables[0].Rows[j]["prior_productionCPH"];
                        dr["high_precisionCPH"] = ds.Tables[0].Rows[j]["high_precisionCPH"];
                        dr["special"] = ds.Tables[0].Rows[j]["special"];
                    }

                    module_Head_TheoryCPH_Table.Rows.Add(dr);

                    da.Dispose();
                    command.Dispose();
                    

                    //得到的module_head_TupleList是代号，需要进行处理
                    line2_module_head_cycletime_qty_TupleList.Add(new Tuple<string, string, string, string>(
                        moduleType,
                        head,
                        truUnitNode[i].SelectSingleNode("./CycleTime").InnerText,
                        truUnitNode[i].SelectSingleNode("./Qty").InnerText
                        )
                        );

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            liteConn.Close();
            return line2_module_head_cycletime_qty_TupleList;
        }


        #region //生成layoutDataTable
        /// <summary>
        /// 获取cycletime\CPH\LayouDT\FeederNozzleDT
        /// </summary>
        /// <param name="productionMode">生产模式</param>
        /// <param name="line2_module_head_cycletime_qty_TupleList">lin2的统计的tuple</param>
        /// <param name="feeder_Statistics">deeder统计</param>
        /// <param name="nozzle_Statistics">nozzle的统计</param>
        /// <returns>
        /// cycletime\CPH\LayouDT\FeederNozzleDT
        /// </returns>
        public Tuple<double, int, DataTable, DataTable> genLayoutEndDT(StreamReader srRTFFile,string productionMode, 
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList,
            Dictionary<string, int> feeder_Statistics, Dictionary<string, int> nozzle_Statistics)
        {
            string[] line1_cycletimes = null;//一轨cycletime
            if (srRTFFile != null)//有一轨数据//Line1线的cycletime从 生成报告的rtf文件中获取
            {
                #region //解析rtf文件，获取line1的cycletime列表
                line1_cycletimes = ParseRTFCycletime(srRTFFile);
                //判断返回值，如果length为0，则代表rtf有问题
                if (line1_cycletimes == null)
                {
                    return null;
                }
                #endregion
            }

            #region 获取生成 layoutDT和feederNozzleDT的数据，totalfeederNozzleDT无
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
            DataTable layoutDT = new DataTable();
            //小表格
            DataTable feederNozzleDT = new DataTable();

            //初始化各个栏
            layoutDT.Columns.Add("Prouction \n Mode");
            layoutDT.Columns.Add("Module \n No.");
            layoutDT.Columns.Add("Module");
            layoutDT.Columns.Add("Head Type");
            layoutDT.Columns.Add("Lane1 \nCycle \nTime");
            layoutDT.Columns.Add("Lane2 \nCycle \nTime");
            layoutDT.Columns.Add("Cycle \nTime");
            layoutDT.Columns.Add("Qty");
            layoutDT.Columns.Add("Avg");
            layoutDT.Columns.Add("CPH");

            feederNozzleDT.Columns.Add("Feeder");
            feederNozzleDT.Columns.Add("Qty Feeder");
            //feederNozzleDT.Columns.Add("Head  Type");
            feederNozzleDT.Columns.Add("Nozzle");
            feederNozzleDT.Columns.Add("QTy  Nozzle");

            //   List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList
            //模组数
            int Module_count = line2_module_head_cycletime_qty_TupleList.Count;
            //feederNozzleDTRowCount、last_layoutExpressDTRowCount。DataTable的行数
            int feederNozzleDTRowCount = nozzle_Statistics.Count > feeder_Statistics.Count ?
                nozzle_Statistics.Count : feeder_Statistics.Count;
            int last_layoutExpressDTRowCount = feederNozzleDTRowCount > Module_count ? feederNozzleDTRowCount : Module_count;
            //循环获取各个单元格的值
            for (int j = 0; j < last_layoutExpressDTRowCount; j++)
            {
                #region layoutDT
                DataRow dr = layoutDT.NewRow();
                if (j < Module_count)//模组的layout
                {
                    dr["Prouction \n Mode"] = productionMode;
                    dr["Module \n No."] = j;
                    dr["Module"] = line2_module_head_cycletime_qty_TupleList[j].Item1;
                    dr["Head Type"] = line2_module_head_cycletime_qty_TupleList[j].Item2;
                    dr["Lane2 \nCycle \nTime"] = line2_module_head_cycletime_qty_TupleList[j].Item3;
                    if (line1_cycletimes != null)
                    {
                        dr["Lane1 \nCycle \nTime"] = line1_cycletimes[j];
                        dr["Cycle \nTime"] = string.Format("{0:0.00}", (System.Convert.ToDouble(line2_module_head_cycletime_qty_TupleList[j].Item3) +
                            System.Convert.ToDouble(line1_cycletimes[j])) / 2);
                    }
                    else
                    {
                        dr["Cycle \nTime"] = line2_module_head_cycletime_qty_TupleList[j].Item3;
                    }
                    dr["Qty"] = line2_module_head_cycletime_qty_TupleList[j].Item4;
                    //line2 每个模组的module head cycletime、Qty--->Avg(cycletime/Qty)、CPH(Qty*3600/cyctime)
                    //计算avg
                    dr["Avg"] = string.Format("{0:0.0000}", Convert.ToDouble(dr["Cycle \nTime"]) / Convert.ToDouble(dr["Qty"]));
                    dr["CPH"] = string.Format("{0:0}", Convert.ToDouble(dr["Qty"]) * 3600 / Convert.ToDouble(dr["Cycle \nTime"]));

                }
                //添加全部行，空行也添加，为了下面的插入最后一行
                layoutDT.Rows.Add(dr);
                #endregion

                //下面feederNozzleDT 直接赋值需要存在DataRow，创建dr
                DataRow drfeederNozzleDT = feederNozzleDT.NewRow();
                feederNozzleDT.Rows.Add(drfeederNozzleDT);
            }
            #region //计算layoutdt最后的一行的值

            DataRow dr2 = layoutDT.NewRow();
            double Line1Cycletime_temp = 0.00;
            double Line2Cycletime_temp = 0.00,
                cycletime_temp = 0.00;
            int qty_temp = 0;
            foreach (DataRow dr in layoutDT.Rows)
            {
                //System.InvalidCastException:“对象不能从 DBNull 转换为其他类型。”
                //空行跳过,判断值是否为DBNull
                if (dr[0].Equals(DBNull.Value))
                {
                    continue;
                }

                if (line1_cycletimes != null)
                {
                    if (Convert.ToDouble(dr[4]) > Line1Cycletime_temp)
                    {
                        Line1Cycletime_temp = Convert.ToDouble(dr[4]);
                    }
                }
                if (Convert.ToDouble(dr[5]) > Line2Cycletime_temp)
                {
                    Line2Cycletime_temp = Convert.ToDouble(dr[5]);
                }


                if (Convert.ToDouble(dr[6]) > cycletime_temp)
                {
                    cycletime_temp = Convert.ToDouble(dr[6]);
                }

                qty_temp += Convert.ToInt32(dr[7]);
            }
            dr2[0] = "Total";
            if (line1_cycletimes != null)
            {
                dr2[4] = Line1Cycletime_temp;
            }
            dr2[5] = Line2Cycletime_temp;
            dr2[6] = cycletime_temp;
            dr2[7] = qty_temp; ;
            dr2[8] = string.Format("{0:0.00}", cycletime_temp * Module_count / qty_temp);
            int CPH = Convert.ToInt32(qty_temp * 3600 / cycletime_temp);
            dr2[9] = CPH;

            //在指定行  last_layoutExpressDTRowCount+1 插入DataRow
            //直接插入指定行，如果指定行前面行不存在，则插入已有行的下面
            layoutDT.Rows.InsertAt(dr2, last_layoutExpressDTRowCount + 1);
            #endregion

            #region //feederNozzleDT
            int k = 0;//在k行直接插入数据
            foreach (var feeder in feeder_Statistics)
            {
                feederNozzleDT.Rows[k]["Feeder"] = feeder.Key;
                feederNozzleDT.Rows[k]["Qty Feeder"] = feeder.Value;
                k++;
            }
            k = 0;//重置，k行插入数据
            foreach (var nozzle in nozzle_Statistics)
            {
                feederNozzleDT.Rows[k]["Nozzle"] = nozzle.Key;
                feederNozzleDT.Rows[k]["QTy  Nozzle"] = nozzle.Value;
                k++;
            }
            #endregion
            ////head type 省略
            //if (j < moduleDT.Rows.Count)
            //{
            //    dr["Head  Type"] = moduleDT.Rows[j]["Head Type"];
            //}

            return new Tuple<double, int, DataTable, DataTable>(cycletime_temp, CPH, layoutDT, feederNozzleDT);
            #endregion
        }
        #endregion


        #region //解析rtf文件，获取line1的cycletime列表
        public string[] ParseRTFCycletime(StreamReader srRTFFile)
        {
            if (srRTFFile==null)
            {
                return null;
            }
            string line = srRTFFile.ReadLine();
            int lineNum = 1; //读取的line行
            int line_CycleTime = 1000;// 读取的line1 cycletime开始的行,用以标记开始，初始设置为1000
            List<string> line1_Cycletime_list = new List<string>();//存放line1 cycletime的list

            bool isCanRead = false;
            while (srRTFFile.EndOfStream != true)
            {
                if (line.Contains("(Lane 1)"))
                {
                    isCanRead = true;
                }
                if (isCanRead)
                {
                    if (line.Contains("CycleTime"))
                    {
                        line_CycleTime = lineNum;
                    }
                    if (line_CycleTime + 2 <= lineNum)//开始解析和读取cycletime的数据
                    {
                        if (line.Contains("-----------"))
                        {//到达line1结尾
                            break;
                        }
                        string[] moduleInfo = line.Split('|');
                        line1_Cycletime_list.Add(moduleInfo[1].Trim());
                    }
                }
                line = srRTFFile.ReadLine();
                lineNum++;
            }
            if (line1_Cycletime_list.Count == 0)
            {
                MessageBox.Show("请确认生产报告的rtf文档正确且为被修改！");
                return null;
            }
            return line1_Cycletime_list.ToArray();
        }
        #endregion

        #region //窗体传值，获取ProductionModel---Dual，Single、double
        public void GetProductionModel()
        {
            ////获取productionModel
            ////productionModleFrom pMFrom = new productionModleFrom(SelectProductionMode);
            ////pMFrom.Show();

            //productionModleFrom pMFrom = new productionModleFrom();
            //pMFrom.ShowDialog(this);//将窗体显示为当前主窗体所有者的模式对话框，productionModleFrom的所有者是当前控件。
            ////while (productionMode==null)
            ////{
            ////    if (productionMode != null)
            ////    {
            ////        break;
            ////    }
            ////}
            productionModleFrom pMFrom = new productionModleFrom();
            if (pMFrom.ShowDialog() == DialogResult.OK)//这样的语句是合法的：DialogResult f = pMFrom.ShowDialog();
            {
                productionMode = pMFrom.production;
            }
        }
        #endregion

        #region //选择生产模式,利用委托传值。未实现
        ////创建委托方法
        //public void SelectProductionMode(string production)
        //{
        //    //groupBox1.Show();
        //    //foreach (Control c in groupBox1.Controls)
        //    //{
        //    //    if ((c is RadioButton) && (c as RadioButton).Checked)
        //    //    {
        //    //        return c.Text;
        //    //    }
        //    //}
        //    productionMode = production;
        //}
        #endregion


        private void GenExcleReport_Btn_Click(object sender, EventArgs e)
        {
            if (summaryandLayout_ComprehensiveList.Count == 0)
            {
                MessageBox.Show("未选择任何可以生成报告的文件夹！");
                return;
            }


            //summaryandLayout_ComprehensiveList 写入到Excel
            if (summaryandLayout_ComprehensiveList.Count > 0)
            {
                //调用保存对话框
                string excelFileName = ComprehensiveStaticClass.SaveExcleDialogShow();
                if (excelFileName == null)
                {
                    return;
                }
                ComprehensiveStaticClass.genExcelfromBaseComprehensiveList(summaryandLayout_ComprehensiveList, excelFileName);

                //保存完后清空列表 summaryandLayout_ComprehensiveList
                //不能设置为null，导出后，再次打开目录 生成处理时summaryandLayout_ComprehensiveList.add添加报 null【未将对象引用设置到对象的实例】
                //summaryandLayout_ComprehensiveList = null;
                summaryandLayout_ComprehensiveList.Clear();//移除所有元素
                //释放资源，未实现

            }
            //将summaryandLayout_ComprehensiveList 列表中的
            //summaryandLayout_Comprehensive按同一机种不同配置、不同配置同一机种写入到Excel
            //if (theEndSummaryandLayout!=null)
            //{
            //    //保存所需信息到Excel
            //    ComprehensiveSaticClass.genExcelfromBaseComprehensive(excelFileName, theEndSummaryandLayout);
            //    //生成成功后重置所有变量的状态
            //    theEndSummaryandLayout = null;
            //}


        }

        private void XmlReport_Load(object sender, EventArgs e)
        {
            //groupBox1.Hide();
        }
    }
}
