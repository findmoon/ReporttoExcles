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
using System.Text.RegularExpressions;

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

        ////模组代码对应的模组类型的 字典
        //Dictionary<string, string> moduleCodeDict = new Dictionary<string, string>()
        //{
        //    {"601","M3III" },
        //    {"602","M6III" }
        //};
        ////Head代码对应的head类型的 字典
        //Dictionary<string, string> headCodeDict = new Dictionary<string, string>()
        //{
        //    {"15","H24" },
        //    {"18","H04SF" }
        //};
                
        //public string ProductionMode
        //{
        //    set
        //    {
        //        productionMode = value;
        //    }
        //}

        //最后包含summary和layout(layouInfoDT feederNozzleDT)的信息的list，实现可以打开多个export文件夹
        List<BaseComprehensive> summaryandLayout_ComprehensiveList = new List<BaseComprehensive>();

        
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

                        try
                        {
                            //创建文件流
                            FileStream fs = new FileStream(files[i], FileMode.Open, FileAccess.Read, FileShare.Read);
                            //读取流
                            srRTFFile = new StreamReader(fs);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                      
                        
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
            //判断xml文件是否缺失
            if (xml_TimingReportHeaderUnit==null|| xml_FeederReportUnit==null || 
                xml_NozzleChangerReportUnit==null || xml_PartReportHead==null || 
                xml_PartReportUnit==null || xml_TimingReportHeader==null||
                xml_TimingReportUnitNxt==null)
            {
                MessageBox.Show("缺少转换所需的导出的xml文件，请确认后重试！");
                
                return;
            }
            //生成最后的 BaseComprehensive List
            //中间出问题，返回null，则不 执行下面的操作
            //只生成 返回 BaseComprehensive
            BaseComprehensive theEndSummaryandLayout = GenBaseComprehensive(xml_TimingReportHeaderUnit,xml_TimingReportUnitNxt,xml_TimingReportHeader,
            xml_PartReportUnit,xml_PartReportHead,xml_NozzleChangerReportUnit,xml_FeederReportUnit,srRTFFile);

            if (theEndSummaryandLayout == null)
            {
                return;
            }
            
            //添加到SummaryandLayout_ComprehensiveList列表
            summaryandLayout_ComprehensiveList.Add(theEndSummaryandLayout);
            //使用全局变量，因为多处需要判断全局变量的值，因此如果选择一个文件夹后处理完所有的过程后，需要重置所有变量，
            //否则导致下一次选择判断的混乱。

            //将 BaseComprehensive List中的BaseComprehensive添加到显示表格中
            //每打开一个文件夹生成一个tabcontrol的tabpage
            AddSingle_EvaluationToTabControl(BaseComprehensiveToSingle_EvaluationPanel(
                summaryandLayout_ComprehensiveList[summaryandLayout_ComprehensiveList.Count - 1]), this.evaluation_TabControl);

        }
        #endregion

        // 每次只会打开一个文件夹，生成Single_EvaluationPanel 也只有一个
        public Single_EvaluationPanelControl BaseComprehensiveToSingle_EvaluationPanel(BaseComprehensive baseComprehensive)
        {
            //baseComprehensives 的layoutDT和feedernozzleDT合在一块
            Single_EvaluationPanelControl single_EvaluationPanelControl= new Single_EvaluationPanelControl(baseComprehensive);
            //还是有右侧的空白
            single_EvaluationPanelControl.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Left);
            //single_EvaluationPanelControl.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            return single_EvaluationPanelControl;
        }

        /// <summary>
        /// 添加单个的Single_EvaluationPanel到tabControl
        /// </summary>
        /// <param name="single_Evaluation">Single_EvaluationPanel对象</param>
        /// <param name="tabControl">TabControl插入的TabControl</param>
        /// index"插入的位置，默认最后一个
        public void AddSingle_EvaluationToTabControl(Single_EvaluationPanelControl single_Evaluation,TabControl tabControl)
        {
            //插入tab，并设置key唯一标识，tab页的标题
            tabControl.TabPages.Add(Convert.ToString(single_Evaluation.TimeStamp), single_Evaluation.JobNmae);
            //获取tab
            tabControl.SelectTab(tabControl.TabPages.Count - 1);
            //tabControl.SelectedTab.Margin= 
            tabControl.SelectedTab.Margin = new Padding(0);
            tabControl.SelectedTab.Padding = new Padding(0);

            //每次生成添加后，要重新调整Single_EvaluationPanel的宽度，否则右侧会有大量空白
            single_Evaluation.Width = tabControl.SelectedTab.Width;
            single_Evaluation.Height = tabControl.SelectedTab.Height;
            single_Evaluation.Parent = tabControl.SelectedTab;
            single_Evaluation.Show();
            
        }

        #region //生成 summaryandLayout_ComprehensiveList  theEndSummaryandLayout
        //生成所有信息 BaseComprehensive class 由summaryinfo和layoutinfo组成
        public BaseComprehensive GenBaseComprehensive(XmlDocument xml_TimingReportHeaderUnit,XmlDocument xml_TimingReportUnitNxt,
            XmlDocument xml_TimingReportHeader,XmlDocument xml_PartReportUnit,XmlDocument xml_PartReportHead,
            XmlDocument xml_NozzleChangerReportUnit,XmlDocument xml_FeederReportUnit,StreamReader srRTFFile)
        {
            
            string jobName = xml_TimingReportHeader.GetElementsByTagName("JobName")[1].InnerText;
            string t_or_b = xml_TimingReportHeader.GetElementsByTagName("BoardSide")[1].InnerText;
            string machine_name = xml_TimingReportHeader.GetElementsByTagName("McName")[1].InnerText;
            //模组总数
            int ModuleCount = xml_TimingReportUnitNxt.GetElementsByTagName("Module").Count - 1;

            XmlNodeList boardNodeList = xml_PartReportUnit.GetElementsByTagName("seqBrdNum");
            int boardQty = GetBoardQty(boardNodeList);

            //使用元组存放module head cycletime  Qt
            //获取 line2_module_head_cycletime_qty_TupleList 元组列表，使用sqlite
            //调用方法
            //module、head、avgCPH、prior_productionCPH	、high_precisionCPH做成一个小DataTable输出.
            //使用out参数
            //DataTable module_Head_TheoryCPH_Table = null;
            //List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList =
            //     Getline2_module_head_cycletime_qty_TupleList(xml_TimingReportUnitNxt,out module_Head_TheoryCPH_Table);
            List<Module_Head_Cph_Struct> module_Head_TheoryCPH_Struct_List = null;
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList =
                Getline2_module_head_cycletime_qty_Tuple_StructList(xml_TimingReportUnitNxt,out module_Head_TheoryCPH_Struct_List);

            //if (module_Head_TheoryCPH_Table == null)
            //{
            //    MessageBox.Show("获取CPH出现了严重错误，请联系作者！");
            //    return ;
            //}

            if (line2_module_head_cycletime_qty_TupleList == null)
            {
                MessageBox.Show("获取CPH出现了严重错误，请联系作者！");
                return null;
            }
            
            #region 统计模组信息，获取字典
            //获取Module的字符串
            string[] allModuleString=new string[line2_module_head_cycletime_qty_TupleList.Count];
            for (int j = 0; j < allModuleString.Length; j++)
            {
                allModuleString[j] = line2_module_head_cycletime_qty_TupleList[j].Item1;
            }
            //获取模组的统计字典
            Dictionary<string, int> module_StatisticsDict =
                ComprehensiveStaticClass.GenModuleStasticsDictFromModuleStrings(allModuleString);

            #endregion

            #region machineKid
            //依据M3III|M6III，判断是NXTIII还是其他，生成machinekind
            string machineKind = "";
            if (module_StatisticsDict.Keys.Contains<string>("M3III") || module_StatisticsDict.Keys.Contains<string>("M6III"))
            {
                machineKind += "NXTIII";
            }
            if (module_StatisticsDict.Keys.Contains<string>("M3II") || module_StatisticsDict.Keys.Contains<string>("M6II"))
            {
                machineKind += "NXTII";
            }
            if (module_StatisticsDict.Keys.Contains<string>("M3S") || module_StatisticsDict.Keys.Contains<string>("M6S") ||
                module_StatisticsDict.Keys.Contains<string>("M3") || module_StatisticsDict.Keys.Contains<string>("M6"))
            {
                machineKind += "NXT";
            }
            #endregion

            #region //Base统计字典
            //获取baseStaxticsDict字典
            Dictionary<string, int> base_StatisticsDict = GetBaseStasticsDict(module_StatisticsDict);
            #endregion
            #region//拼接Line字符串
            //获取拼接的Line字符串
            string Line = ComprehensiveStaticClass.GenLineStringFromModuleStasticsBaseStastics(
                module_StatisticsDict, base_StatisticsDict, machineKind);
            #endregion

            #region//获取组合"Head Type"的信息
            //所有头的字符串
            string[] allHeadString = new string[line2_module_head_cycletime_qty_TupleList.Count];
            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                allHeadString[j] = line2_module_head_cycletime_qty_TupleList[j].Item2;
            }

            //获取拼接的Head_Type
            string Head_Type = ComprehensiveStaticClass.GenHeadTypeFormHeadStrings(allHeadString);
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
            string remark = ComprehensiveStaticClass.GenRemarkFromModuleBaseStasticsDict(
                module_StatisticsDict, base_StatisticsDict);
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
                    if (feederName.InnerText=="")
                    {
                        continue;
                    }
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
            //存放生产模式的变量
            string productionMode = null;
            productionMode = GetProductionModel();

            //获取 genLayoutEndDT，最后的cycletime cph theory_cph layoutDT、feedernozzleDT
            Tuple<double, int, DataTable, DataTable> theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple = 
                GetEndCycletimeCPHLayoutDTFeederNozzleDT(srRTFFile,productionMode,
                line2_module_head_cycletime_qty_TupleList, feeder_Statistics, nozzle_Statistics);
            //判断是否获取到 
            if (theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple == null)
            {
                return null;
            }

            /// Output Board/hour  -----> board qty *3600/cycletime
            // Output Board/Day(22H)  ---->(Output Board/hour)*22
            //CPH Rate   ----> CPH/理论CPH
            // CPP  链接报价单，未实现
            //以上由函数GetTheEndSummaryInfoDT实现
                                    
            #region //pictureData放在前面，未生成则不进行下面对象的创建
            //判断是否获取成功图片的byte[]
            //所有模组type的string。获取最后的图片 byte
            byte[] pictureDataByte = ComprehensiveStaticClass.GenJobsPicture(allModuleString, module_StatisticsDict);
            if (pictureDataByte == null)
            {
                return null;
            }
            #endregion


            //获取最后的综合的 SummaryandLayout
            //最后包含summaryInfo和layouInfoDT feederNozzleDT的信息
            BaseComprehensive theEndSummaryAndLayout = new BaseComprehensive(ComprehensiveStaticClass.GetTimeStamp());
            theEndSummaryAndLayout.T_or_B = t_or_b;
            theEndSummaryAndLayout.ModuleCount = ModuleCount;
            theEndSummaryAndLayout.MachineKind = machineKind;
            theEndSummaryAndLayout.Jobname = jobName;
            //获取CPHRate
            //dr["module"] 
            //dr["Head Type"] 
            //dr["avgCPH"]
            //dr["prior_productionCPH"]
            //dr["high_precisionCPH"]
            //dr["special"]
            int theoryCPH = 0;
            //for (int i = 0; i < module_Head_TheoryCPH_Table.Rows.Count; i++)
            //{
            //    //判断高速生产模式cph是否为 空

            //    if (module_Head_TheoryCPH_Table.Rows[i]["prior_productionCPH"].ToString()!="0")
            //    {
            //        theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["prior_productionCPH"].ToString());
            //    }
            //    else if (module_Head_TheoryCPH_Table.Rows[i]["avgCPH"].ToString() != "0")
            //    {
            //        theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["avgCPH"].ToString());
            //    }
            //    else if (module_Head_TheoryCPH_Table.Rows[i]["high_precisionCPH"].ToString() != "0")
            //    {
            //        theoryCPH += System.Convert.ToInt32(module_Head_TheoryCPH_Table.Rows[i]["high_precisionCPH"].ToString());
            //    }

            //}
            //string cphRate = string.Format("{0:P}", System.Convert.ToDouble(layoutEndInfo_tuple.Item2) / System.Convert.ToDouble(theoryCPH));
            //machine_name

            /*
             * public struct module_head_cph_struct
        {
            public string[] moduleTrayTypestring;

            public List<string> headTypeString_List;
            public List<string[]> CPH_strings_List;
        }
             */

            //使用结构计算theoryCPH
            for (int j = 0; j < module_Head_TheoryCPH_Struct_List.Count; j++)
            {
                //"高速生产模式-1234"需要解析
                theoryCPH += Convert.ToInt32(module_Head_TheoryCPH_Struct_List[j].CPH_strings_List[0][0].Split('-')[1]);
            }
            string cphRate = string.Format("{0:P}", Convert.ToDouble(
                theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple.Item2) / Convert.ToDouble(theoryCPH));

            //获取最后的SummaryinfoDT
            theEndSummaryAndLayout.GetTheEndSummaryInfoDT(expressSummayDataTable,theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple.Item1, 
                theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple.Item2,cphRate);

            theEndSummaryAndLayout.PictureDataByte = pictureDataByte;

            theEndSummaryAndLayout.layoutDT = theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple.Item3;
            theEndSummaryAndLayout.feederNozzleDT = theEndCycletimeCPHLayoutDTFeederNozzleDT_Tuple.Item4;
            theEndSummaryAndLayout.AllModuleType = allModuleString;
            theEndSummaryAndLayout.MachineKind = machineKind;

            //所有头Head type的string
            string[] allHeadTypeString = new string[line2_module_head_cycletime_qty_TupleList.Count];
            for (int j = 0; j < line2_module_head_cycletime_qty_TupleList.Count; j++)
            {
                allHeadTypeString[j] = line2_module_head_cycletime_qty_TupleList[j].Item2;
            }
            theEndSummaryAndLayout.AllHeadType = allHeadTypeString;
            theEndSummaryAndLayout.module_Head_Cph_Structs_List = module_Head_TheoryCPH_Struct_List;
            theEndSummaryAndLayout.base_StatisticsDict = base_StatisticsDict;
            theEndSummaryAndLayout.module_StatisticsDict = module_StatisticsDict;

            return theEndSummaryAndLayout;
            
        }

        #endregion

        private Dictionary<string, int> GetBaseStasticsDict(Dictionary<string, int> module_Statistics)
        {
            //计算base信息
            //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
            //base尽可能少的分配
            //转换为单M的base 的总数
            int baseCount_M_TotalM3 = 0;
            //单独统计M3、M6个数
            int Count_M3 = 0;
            int Count_M6 = 0;
            foreach (var moduleKind in module_Statistics)
            {
                if (moduleKind.Key.Substring(0, 2) == "M3")
                {
                    baseCount_M_TotalM3 += moduleKind.Value;
                    Count_M3 += moduleKind.Value;
                }
                if (moduleKind.Key.Substring(0, 2) == "M6")
                {
                    baseCount_M_TotalM3 += moduleKind.Value * 2;
                    Count_M6 += moduleKind.Value;
                }
            }
            int baseCount_4M = (baseCount_M_TotalM3 / 4);
            int baseCount_2M = 0;
            if ((baseCount_M_TotalM3 % 4) != 0)
            {
                baseCount_2M = 1;
            }

            //字典，存放2MBase 4MBase 的数量
            return new Dictionary<string, int>() {
                {"4MBASE",baseCount_4M },
                {"2MBASE",baseCount_2M },
                };
        }

        /// <summary>
        /// 获取BoardQty
        /// </summary>
        /// <param name="boardNodeList"></param>
        /// <returns></returns>
        private int GetBoardQty(XmlNodeList boardNodeList)
        {
            int boardQty = Int32.Parse(boardNodeList[1].InnerText);
            for (int i = 1; i < boardNodeList.Count; i++)
            {
                if (boardQty < Int32.Parse(boardNodeList[i].InnerText))
                {
                    boardQty = Int32.Parse(boardNodeList[i].InnerText);
                }
            }
            return boardQty;
        }

        #region //旧的 XmlDocument 中获取module head cycletime qty和包含module_Head_TheoryCPH的DataTable
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
                        //此处错误，ds.Tables[0] select查询出来的结果是无顺序的，第一个并不一定是非特殊头
                        if (j > 0)//先跳过H24 S G head 的处理
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

        #endregion

        #region //XmlDocument 中获取module head cycletime qty
        //和包含module_Head_TheoryCPH所有可能性的List<module_head_cph_struct>
        public List<Tuple<string, string, string, string>> Getline2_module_head_cycletime_qty_Tuple_StructList(XmlDocument xml_TimingReportUnitNxt,
            out List<Module_Head_Cph_Struct> module_Head_TheoryCPH_Struct_List)
        {
            //使用元组存放module head cycletime  Qty
            //获取 line2_module_head_cycletime_qty_TupleList 元组列表，使用sqlite
            //调用方法
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList =
                new List<Tuple<string, string, string, string>>();

            //获取line对应的module\head\Qty|cycletime信息 
            XmlNodeList truUnitNode = xml_TimingReportUnitNxt.GetElementsByTagName("Unit");


            //初始化 module_Head_TheoryCPH_listTuple_List
            module_Head_TheoryCPH_Struct_List = new List<Module_Head_Cph_Struct>();

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
                    #region //moduleType Head
                    string moduleid = truUnitNode[i].SelectSingleNode("./cfgModuleType1").InnerText;
                    string headid = truUnitNode[i].SelectSingleNode("./cfgHeadType1_1").InnerText;

                    //sqlite的命令对象,查询moduleType
                    SQLiteCommand command = new SQLiteCommand(string.Format("select moduletype,trayinfostring from ModuleTable where moduleid={0};", moduleid),
                        liteConn);
                    if (liteConn.State == ConnectionState.Closed)
                    {
                        liteConn.Open();
                    }
                    SQLiteDataReader dataRead = command.ExecuteReader();
                    //string moduleType = dataRead["moduletype"].ToString();

                    //获取所有的 module /+try 类型
                    List<string> moduleTrayTypeList = new List<string>();
                    while (dataRead.Read())
                    {
                        moduleTrayTypeList.Add(dataRead["moduletype"].ToString());
                        if (!dataRead["trayinfostring"].Equals(DBNull.Value))
                        {
                            string[] moduleTrayType = dataRead["trayinfostring"].ToString().Split('-');
                            for (int j = 0; j < moduleTrayType.Length; j++)
                            {
                                moduleTrayTypeList.Add(moduleTrayTypeList[0]+"-"+moduleTrayType[j]);
                            }
                        }
                    }


                    List<string> headTypeList = new List<string>();
                    //查询headType
                    command = new SQLiteCommand(string.Format("select head from HeadTable where headid={0};", headid),
                        liteConn);
                    if (liteConn.State == ConnectionState.Closed)
                    {
                        liteConn.Open();
                    }

                    headTypeList.Add(command.ExecuteScalar().ToString());
                    #endregion

                    #region //查询CPH
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

                    //判断下dataset
                    if (ds.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show(string.Format("未找到{0}和{1}对应的理论CPH",
                            moduleTrayTypeList[0], headTypeList[0]));
                        return null;
                    }

                    //datarow module_Head_TheoryCPH_Table
                    //DataRow dr = module_Head_TheoryCPH_Table.NewRow();

                    //当前cph module对应的CPH的List
                    List<string[]> CPH_List = new List<string[]>();

                    //存放当前 head module的的三个或2个CPH
                    List<string> cph_string = new List<string>();

                    //模组+tray的变换不是联动
                    //head和CPH的数值 是联动的，同一个head下对应三个cph取值，这个由List<string[]> list下标对应

                    //处理special = null时的数据
                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                    {
                        //处理逻辑不是为null的时候插入，而是为null的时候必须插入第一个。
                        //select获取的数据时无需要，因此要先选出为null赋值
                        //原值为null，获取后为字符串"null",因此判断
                        if (ds.Tables[0].Rows[j]["special"].ToString() == "null")
                        {
                            //判断高速生产模式cph是否为 空
                            //将非 H24G、H24S 头的 对应的 高速、正常、高精度模式记录
                            if (ds.Tables[0].Rows[j]["avgCPH"].ToString() != "0")
                            {
                                cph_string.Add("正常模式-" + ds.Tables[0].Rows[j]["avgCPH"].ToString());
                            }

                            if (ds.Tables[0].Rows[j]["prior_productionCPH"].ToString() != "0")
                            {
                                cph_string.Add("高速生产模式-" + ds.Tables[0].Rows[j]["prior_productionCPH"].ToString());
                            }

                            if (ds.Tables[0].Rows[j]["high_precisionCPH"].ToString() != "0")
                            {
                                cph_string.Add("高精度模式-" + ds.Tables[0].Rows[j]["high_precisionCPH"].ToString());
                            }
                            CPH_List.Add(cph_string.ToArray());
                            cph_string.Clear();
                        }
                    }

                    //处理其他 H24G H24S 的数据：
                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                    {
                        //处理逻辑不是未null的时候插入，而是为null的时候必须插入第一个。
                        //select获取的数据时无需要，因此要先选出为null赋值
                        if (ds.Tables[0].Rows[j]["special"].ToString() != "null")
                        {
                            //添加特殊head
                            headTypeList.Add(ds.Tables[0].Rows[j]["special"].ToString());
                            //判断高速生产模式cph是否为 空
                            //将  头的 对应的 高速、正常、高精度模式记录
                            if (ds.Tables[0].Rows[j]["avgCPH"].ToString() != "0")
                            {
                                cph_string.Add("正常模式-" + ds.Tables[0].Rows[j]["avgCPH"].ToString());
                            }

                            if (ds.Tables[0].Rows[j]["prior_productionCPH"].ToString() != "0")
                            {
                                cph_string.Add("高速生产模式-" + ds.Tables[0].Rows[j]["prior_productionCPH"].ToString());
                            }

                            if (ds.Tables[0].Rows[j]["high_precisionCPH"].ToString() != "0")
                            {
                                cph_string.Add("高精度模式-" + ds.Tables[0].Rows[j]["high_precisionCPH"].ToString());
                            }
                            CPH_List.Add(cph_string.ToArray());
                            cph_string.Clear();
                        }
                    }

                    da.Dispose();
                    command.Dispose();
                    #endregion
                    //判断当前module head对应的 的specia head 和 cph的个数是否对应
                    if (headTypeList.Count != CPH_List.Count)
                    {
                        MessageBox.Show("出现了意外，请暂停使用！");
                        return null;
                    }
                    //当前module head对应的 module_head_cph_struct结构
                    Module_Head_Cph_Struct module_Head_Cph_Struct = new Module_Head_Cph_Struct()
                    {
                        moduleTrayTypestring = moduleTrayTypeList.ToArray(),
                        headTypeString_List = headTypeList,
                        CPH_strings_List = CPH_List
                    };
                    module_Head_TheoryCPH_Struct_List.Add(module_Head_Cph_Struct);


                    //module_head_cycletime_qty  module_head 取值第一个
                    line2_module_head_cycletime_qty_TupleList.Add(new Tuple<string, string, string, string>(
                        module_Head_Cph_Struct.moduleTrayTypestring[0],
                        module_Head_Cph_Struct.headTypeString_List[0],
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
        #endregion

        #region //生成layoutDataTable
        /// <summary>
        /// 获取计算后的cycletime\CPH\LayouDT\FeederNozzleDT
        /// </summary>
        /// <param name="productionMode">生产模式</param>
        /// <param name="line2_module_head_cycletime_qty_TupleList">lin2的统计的tuple</param>
        /// <param name="feeder_Statistics">deeder统计</param>
        /// <param name="nozzle_Statistics">nozzle的统计</param>
        /// <returns>
        /// cycletime\CPH\LayouDT\FeederNozzleDT
        /// </returns>
        public Tuple<double, int, DataTable, DataTable> GetEndCycletimeCPHLayoutDTFeederNozzleDT(StreamReader srRTFFile,string productionMode, 
            List<Tuple<string, string, string, string>> line2_module_head_cycletime_qty_TupleList,
            Dictionary<string, int> feeder_Statistics, Dictionary<string, int> nozzle_Statistics)
        {
            Tuple<string[],string[]> lane1_lane2_cycletimeStrings_Tuple = null;//一轨cycletime
            if (srRTFFile != null)//有一轨数据//Line1线的cycletime从 生成报告的rtf文件中获取
            {
                #region //解析rtf文件，获取line1的cycletime列表
                lane1_lane2_cycletimeStrings_Tuple = ParseRTFCycletime(srRTFFile);
                //判断返回值，为null，则代表没有获取到解析的cycletime
                if (lane1_lane2_cycletimeStrings_Tuple == null)
                {
                    return null;
                }
                #endregion
            }

            //根绝productionMode 判断使用获取到的哪个轨道的cycyletime（获取单双轨道的cycletime）
            //if (productionMode == "Dual-DoubleLane")
            //{

            //}
            ////获取Lane1的cycyletime
            //if (productionMode == "Single-Lane 1")
            //{

            //}
            ////获取lane2的cycletime
            //if (productionMode == "Single-Lane 2")
            //{

            //}
            

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
                    dr["Prouction \n Mode"] = productionMode.Split('-')[0];
                    dr["Module \n No."] = j+1;//模组从1开始
                    dr["Module"] = line2_module_head_cycletime_qty_TupleList[j].Item1;
                    dr["Head Type"] = line2_module_head_cycletime_qty_TupleList[j].Item2;
                    if (productionMode == "Dual-DoubleLane")
                    {
                        dr["Lane1 \nCycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item1[j];
                        dr["Lane2 \nCycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item2[j];
                        
                        dr["Cycle \nTime"] = string.Format("{0:0.00}", (Convert.ToDouble(lane1_lane2_cycletimeStrings_Tuple.Item1[j]) +
                            Convert.ToDouble(lane1_lane2_cycletimeStrings_Tuple.Item2[j])) / 2);
                    }
                    //获取Lane1的cycyletime
                    if (productionMode == "Single-Lane 1")
                    {
                        dr["Lane1 \nCycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item1[j];                        
                        dr["Cycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item1[j];
                    }
                    //获取lane2的cycletime
                    if (productionMode == "Single-Lane 2")
                    {
                        dr["Lane2 \nCycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item2[j];
                        dr["Cycle \nTime"] = lane1_lane2_cycletimeStrings_Tuple.Item2[j];
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

                if (productionMode == "Dual-DoubleLane")
                {
                    if (Convert.ToDouble(dr["Lane1 \nCycle \nTime"]) > Line1Cycletime_temp)
                    {
                        Line1Cycletime_temp = Convert.ToDouble(dr["Lane1 \nCycle \nTime"]);
                    }
                    if (Convert.ToDouble(dr["Lane2 \nCycle \nTime"]) > Line2Cycletime_temp)
                    {
                        Line2Cycletime_temp = Convert.ToDouble(dr["Lane2 \nCycle \nTime"]);
                    }
                    
                }
                //获取Lane1的cycyletime
                if (productionMode == "Single-Lane 1")
                {
                    if (Convert.ToDouble(dr["Lane1 \nCycle \nTime"]) > Line1Cycletime_temp)
                    {
                        Line1Cycletime_temp = Convert.ToDouble(dr["Lane1 \nCycle \nTime"]);
                    }
                }
                //获取lane2的cycletime
                if (productionMode == "Single-Lane 2")
                {
                    if (Convert.ToDouble(dr["Lane2 \nCycle \nTime"]) > Line2Cycletime_temp)
                    {
                        Line2Cycletime_temp = Convert.ToDouble(dr["Lane2 \nCycle \nTime"]);
                    }
                }
                
                if (Convert.ToDouble(dr[6]) > cycletime_temp)
                {
                    cycletime_temp = Convert.ToDouble(dr[6]);
                }

                qty_temp += Convert.ToInt32(dr[7]);
            }
            dr2[0] = "Total";
            
            //一二轨赋值
            if (productionMode == "Dual-DoubleLane")
            {
                dr2[4] = Line1Cycletime_temp;
                dr2[5] = Line2Cycletime_temp;
            }
            //一轨
            if (productionMode == "Single-Lane 1")
            {
                dr2[4] = Line1Cycletime_temp;
            }
            //二轨
            if (productionMode == "Single-Lane 2")
            {
                dr2[5] = Line2Cycletime_temp;
            }
            
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

        #region //解析rtf文件，获取line1的cyctime 字符串数组
        public Tuple<string[],string[]> ParseRTFCycletime(StreamReader srRTFFile)
        {
            if (srRTFFile == null)
            {
                return null;
            }
            //创建读取流
            //StreamReader sr = new StreamReader(fs, Encoding.Default);

            //最后lane1、lane2对应的cycle列表
            List<string> lane1_Cycletime_list = new List<string>();//存放lane1 cycletime的list
            List<string> lane2_Cycletime_list = new List<string>();//存放lane2 cycletime的list
            //获取文件 所有的内容
            string rtfText = srRTFFile.ReadToEnd();

            //Lane 1 \ 2 最后匹配拆分的 包含cycletime的集合
            MatchCollection lane1CycleMatchCollection = null;
            MatchCollection lane2CycleMatchCollection = null;

            //写入文件无法显示中文。原字符创中也无法显示中文。
            //using (StreamWriter fileWriter = new StreamWriter("./filWriter.txt", true, Encoding.GetEncoding("gb2312")))
            //{
            //    fileWriter.Write(line);
            //}

            
            //获取lane 1的报告部分
            //. 匹配换行符
            MatchCollection lane1matchCollection = Regex.Matches(rtfText, "Lane 1(.*?)MAX", RegexOptions.Singleline);
            MatchCollection lane2matchCollection = Regex.Matches(rtfText, "Lane 2(.*?)MAX", RegexOptions.Singleline);
            string lane1AllString = lane1matchCollection[0].Groups[1].Value;
            string lane2AllString = lane2matchCollection[0].Groups[1].Value;
            //textBox1.Text = lane1AllString;

            #region //获取Lane 1
            if (lane1AllString.Split('\n').Length > 1)//可以读取多行时，或存在多行时
            {
                //匹配lian 的包含数字，姐cycletime部分
                string regPattrenString = @"\\par(.*?\d+?.*?)\n";
                lane1CycleMatchCollection = Regex.Matches(lane1AllString, regPattrenString, RegexOptions.Singleline);

            }
            else
            {
                //匹配lian 的包含数字，即cycletime部分
                string lane1regPattrenString = @"{\\fs18 \\kerning2 \\loch \\af1 \\hich \\af1 \\dbch \\f1 \\lang2052 \\langnp2052 \\cf1(.*?\d*?.*?)}";
                lane1CycleMatchCollection = Regex.Matches(lane1AllString, lane1regPattrenString, RegexOptions.Singleline);

            }
            //匹配获取成功，匹配项数目至少为1
            if (lane1CycleMatchCollection.Count>0)
            {
                //解析获取的每一组，以数字开头的为模组数 对应的数据
                string lane1moduleRegPatternString = @"^\d.*";
                for (int i = 0; i < lane1CycleMatchCollection.Count; i++)
                {
                    //逐个解析正则出来的每一组，包含

                    //当前组的内容
                    string line = lane1CycleMatchCollection[i].Groups[1].Value.Trim();

                    Match moduleMatch = Regex.Match(line, lane1moduleRegPatternString);
                    //匹配成功
                    if (moduleMatch.Success)
                    {
                        string[] moduleInfo = line.Split('|');

                        lane1_Cycletime_list.Add(moduleInfo[1].Trim());
                    }

                }
            }
            
            #endregion

            #region //获取Lane 2
            if (lane2AllString.Split('\n').Length > 1)//可以读取多行时，或存在多行时
            {
                //匹配lian 的包含数字，姐cycletime部分
                string lane2regPattrenString = @"\\par(.*?\d+?.*?)\n";
                lane2CycleMatchCollection = Regex.Matches(lane2AllString, lane2regPattrenString, RegexOptions.Singleline);

            }
            else
            {
                //匹配lian 的包含数字，即cycletime部分
                string lane2regPattrenString = @"{\\fs18 \\kerning2 \\loch \\af1 \\hich \\af1 \\dbch \\f1 \\lang2052 \\langnp2052 \\cf1(.*?\d*?.*?)}";
                lane2CycleMatchCollection = Regex.Matches(lane2AllString, lane2regPattrenString, RegexOptions.Singleline);

            }
            //匹配获取成功，匹配项数目至少为1
            if (lane2CycleMatchCollection.Count > 0)
            {
                //解析获取的每一组，以数字开头的为模组数 对应的数据
                string lane2moduleRegPatternString = @"^\d.*";
                for (int i = 0; i < lane2CycleMatchCollection.Count; i++)
                {
                    //逐个解析正则出来的每一组，包含

                    //当前组的内容
                    string line = lane2CycleMatchCollection[i].Groups[1].Value.Trim();

                    Match moduleMatch = Regex.Match(line, lane2moduleRegPatternString);
                    //匹配成功
                    if (moduleMatch.Success)
                    {
                        string[] moduleInfo = line.Split('|');

                        lane2_Cycletime_list.Add(moduleInfo[1].Trim());
                    }

                }
            }
            #endregion
            if (lane1_Cycletime_list.Count == 0&&lane2_Cycletime_list.Count==0)
            {
                MessageBox.Show("出现了意外，无法解析rtf文件，或rtf文件中没有Lane1和Lane2的数据");
                //出现问题重置所有情况
                //释放资源
                srRTFFile.Dispose();
                //重置srRTFFile
                srRTFFile = null;
                return null;
            }
            //区分是Single-Lane 1\Single-Lane 2\Dual-DuobleLane
            
            srRTFFile.Dispose();//释放占用的rtf文件
            //重置srRTFFile
            srRTFFile = null;
            return new Tuple<string[], string[]>(lane1_Cycletime_list.ToArray(), lane2_Cycletime_list.ToArray());
                
        }
        #endregion

        #region //解析rtf文件，获取line1的cycletime列表
        public string[] ParseRTFCycletime_Old(StreamReader srRTFFile)
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
                //出现问题重置所有情况
                //释放资源
                srRTFFile.Dispose();
                //重置srRTFFile
                srRTFFile = null;
                return null;
            }
            srRTFFile.Dispose();//释放占用的rtf文件
            return line1_Cycletime_list.ToArray();
        }
        #endregion

        #region //窗体传值，获取ProductionModel---Dual，Single、double
        public string GetProductionModel()
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
            string productionMode=null;
            productionModleFrom pMFrom = new productionModleFrom();
            if (pMFrom.ShowDialog() == DialogResult.OK)//这样的语句是合法的：DialogResult f = pMFrom.ShowDialog();
            {
                productionMode = pMFrom.production;
            }

            return productionMode;
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

        #region /// <summary>
        /// module head CPH最后的对应关系设计为一个字段
        /// </summary>
        //public struct module_head_cph_struct
        //{
        //    public string[] moduleTrayTypestring;

        //    public List<string> headTypeString_List;
        //    public List<string[]> CPH_strings_List;
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
                ComprehensiveStaticClass.GenExcelfromBaseComprehensiveList(summaryandLayout_ComprehensiveList, excelFileName);

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

        private void latestBaseComprehensive_treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
