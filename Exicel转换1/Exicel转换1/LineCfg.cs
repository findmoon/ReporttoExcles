using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace Exicel转换1
{
    public partial class LineCfg : UserControl
    {
        public LineCfg()
        {
            InitializeComponent();
        }

        string fileDAT=null;
        string fileMht=null;

        //最后转化的内容DataTable
        DataTable lineCFGDatatable=new DataTable();
        //初始化DataTable的列
        string[] lineCFGDatatable_Columns = { "FactoryName", "LineName", "ModelName","NxtModuleInfo","ModuleType",
            "Module/Machine Number", "MachineName", "MachineType", "Position", "BaseName", "ServerName", "BaseTypeID", "BaseWidth" };

        //模组对应关系的字典
        Dictionary<int, string> MuduleInt_dct = new Dictionary<int, string>()
        {
            {1,"M3" },
            {2,"M6" }
        };

        //加载后的mht下body元素
        HtmlElementCollection bodyElems;

        private void OpenFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                if (openFileDialog1.FileName.IndexOf(".dat")>0|| openFileDialog1.FileName.IndexOf(".DAT") > 0)
                {
                    fileDAT = openFileDialog1.FileName;
                }
                if (openFileDialog1.FileName.IndexOf(".mht") > 0)
                {
                    fileMht = openFileDialog1.FileName;
                }
                if (fileDAT!=null&& fileMht!=null)
                {
                    MessageBox.Show("请将选择的layout文件导出");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (fileDAT==null)
            {
                MessageBox.Show("未选择dat文件！");
                return;
            }
            if (fileMht==null)
            {
                MessageBox.Show("未选择layout.mht文件！");
                return;
            }

            //初始化lineCFGDatatable
            for (int i = 0; i<lineCFGDatatable_Columns.Length; i++)
            {
                lineCFGDatatable.Columns.Add(lineCFGDatatable_Columns[i]);
            }
            //解析dat文件
            ParseDATFile(fileDAT, lineCFGDatatable);
            //解析mht文件
        }

        public void ParseDATFile(string filePath,DataTable lineCFGDatatable)
        {
            //创建文件流
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            //读取流
            StreamReader sr = new StreamReader(fs);

            //逐行读取
            string readLine = sr.ReadLine();
            //以L行为界，每一次L开头行为对应的一组，中间内容为对应的DataTable一行
            int dataTableRow_num = 1;
            //dr放在外面，由L开头行创建
            DataRow dr = null;

            while (readLine!= "- End")
            {
                //拆分读到的行
                string[] Line_Strings = readLine.Split('\t');
                
                //解析每行
                if (Regex.IsMatch(readLine,@"^L"))//L 开头
                {
                    /*
                    if (dr!=null)//如果有存在行，加到DataTable并清空
                    {
                        lineCFGDatatable.Rows.Add(dr);
                        dr = null;
                    }
                    */

                    dr = lineCFGDatatable.NewRow();
                    dr[0] = Line_Strings[1];
                    dr[1] = Line_Strings[2];
                }
                
                //在L开始的一组下解析
                 
                    if (Regex.IsMatch(readLine, @"^M"))
                    {
                        //ModelName
                        dr[2] = Line_Strings[1];
                        //NxtModelInfo
                        string[] NxtModelInfo_strings=Line_Strings[2].Split();
                        dr[5]= NxtModelInfo_strings.Length;
                        int module_int;//model对应的数字
                        if (Int32.TryParse(NxtModelInfo_strings[0], out module_int))
                        {//能解析为Nxt
                            string NxtModelInfo_string=null;

                            //M6^ M3个数
                            int M3_int=0,M6_int=0;
                            for (int i = 0; i < NxtModelInfo_strings.Length; i++)
                            {
                                try
                                {
                                    module_int=System.Convert.ToInt32(NxtModelInfo_strings[i]);
                                    NxtModelInfo_string += MuduleInt_dct[module_int] + " ";
                                    if (module_int==1)
                                    {
                                        M3_int++;
                                    }
                                    if (module_int == 2)
                                    {
                                        M6_int++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw ex;
                                }
                            }

                            NxtModelInfo_string += "\n" + M3_int + "个M3 " + M6_int + "个M6";
                            dr[3] = NxtModelInfo_string;
                        }
                        else
                        {
                            dr[3] = Line_Strings[2];
                        }
                    }
                    //"FactoryName", "LineName", "ModelName","NxtModuleInfo","ModuleType",
                    //"Module/Machine Number", "MachineName", "MachineType" 7, "Position" 8, 
                    //"BaseName", "ServerName" 10, "BaseTypeID", "BaseWidth"
                    if (Regex.IsMatch(readLine, @"^T"))
                    {
                        dr[6] = Line_Strings[1];
                        dr[7] = Line_Strings[2];
                    }
                    if (Regex.IsMatch(readLine, @"^B"))
                    {
                        dr[8] = Line_Strings[1];
                        dr[9] = Line_Strings[2];
                        dr[10] = Line_Strings[4];
                        dr[11] = Line_Strings[5];
                        dr[12] = Line_Strings[6];
                        //B行开始解析完后，增加一行
                        lineCFGDatatable.Rows.Add(dr);
                        dr = null;
                        dataTableRow_num++;//增加一行
                    }
                readLine = sr.ReadLine();                
            }
        }

        public void LoadedMhtFile(string filePath)
        {
            WebBrowser webBrowser1 = new WebBrowser();
            webBrowser1.Url = new Uri(filePath);
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.Url = webBrowser1.Url;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(WebBrowser_DocumentCompleted);
        }

        void WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;            
            bodyElems = webBrowser.Document.GetElementsByTagName("body");
            //MessageBox.Show(elems.Count.ToString());
            //}
            //加载完成后解析
            ParthMhtElem(bodyElems[0], lineCFGDatatable);
        }
        public void ParthMhtElem(HtmlElement bodyElem, DataTable lineCFGDatatable)
        {
            //获取中间id=Table1
            HtmlElementCollection table1Element = bodyElem.All.GetElementsByName("Table1");
        }


        /*       
       'FactoryName, LineName, ModelName, MachineName, MachineType, Position, BaseName, ServerName, BaseTypeID, BaseWidth
       '将数据展开并复制到一旁,L>1,2，M>1,2,T>1,2,B>1,2,4,5,6
       Dim Ltitle(0 To 2) As String, Mtitle(3) As String, Ttitle(2) As String, Btitle(5) As String


       Ltitle(0) = "FactoryName"
       Ltitle(1) = "LineName"

       Mtitle(0) = "ModelName"
       Mtitle(1) = "NxtModuleInfo"
       Mtitle(2) = "Module/Machine Number"

       Ttitle(0) = "MachineName"
       Ttitle(1) = "MachineType"

       Btitle(0) = "Position"
       Btitle(1) = "BaseName"
       Btitle(2) = "ServerName"
       Btitle(3) = "BaseTypeID"
       Btitle(4) = "BaseWidth"

       Dim condit, max, LK, Lj, MK, Mj, TK, Tj, DK, Dj, SK, Sj, Ek, Ej, BK, Bj, j, i, anotherLstart
       condit = 1
       max = 1
       LK = 1
       MK = 1
       TK = 1
       DK = 1
       SK = 1
       Ek = 1
       BF = 1


       While Cells(condit, 1) <> ""

           Dim MArr() '声明动态数组，动态存放modelinfozhong 2212 等数据
           If Cells(condit, 1) = "L" Then
               LK = max
               anotherLstart = LK
               Lj = 2
               While Cells(condit, Lj) <> ""
                   Cells(1 + LK, 10 + Lj) = Cells(condit, Lj)


                   Lj = Lj + 1
               Wend
               LK = LK + 1
                   If max<LK Then
                       max = LK
                   End If
               condit = condit + 1
           End If

           '本组出现的L行复制结束




               i = condit


               MK = anotherLstart
               TK = anotherLstart
               SK = anotherLstart
               BK = anotherLstart
               DK = anotherLstart
               Ek = anotherLstart
               '每一组L组开始之前，都要对其当前组L开始的行数


               While Cells(i, 1) <> "L" And Cells(i, 1) <> ""

                      If Cells(i, 1) = "M" Then
                           'Mj = Lj '若是中间有缺少或长短不一的列，这个赋值将不准确，因此给取固定长度
                           'Mj = 4
                           Mj = 2 + UBound(Ltitle)

                           'MK = anotherLstart

                               Cells(1 + MK, 10 + Mj) = Cells(i, 2)
                               Dim MType, MEndType, leng, M, z, M3Count, M6Count, BBCount, XCount, haveNXT
                               MType = Cells(i, 3)
                                'MsgBox (MType)MType显示输出

                               haveNXT = True
                               M3Count = 0
                               M6Count = 0
                               BBCount = 0
                               XCount = 0

                               leng = Len(MType) - 1 '单元格中数字字符串的长度
                               ReDim MArr(Len(MType))


                               For z = 0 To leng
                                   MArr(z) = Right(Left(MType, z + 1), 1)

                               Next z

                               'MsgBox (Len(MType))
                               'MsgBox (LBound(MArr))
                               'MsgBox (UBound(MArr))UBound(MArr)指的是数组的上标界限，比数组最大标记多1
                               'MsgBox (MArr(UBound(MArr)))‘但直接取值并不抱错数组越界
                               'MsgBox (MArr(5))'直接超出下标取值显示会显示越界，超出一个只会显示为空

                               For M = LBound(MArr) To(UBound(MArr) - 1)
                                   'MsgBox (MArr(M))使用Ubound（MArr）会得到最后一位为空的位，需要减1才能保证不多取
                                   If MArr(M) = 1 Then
                                       MArr(M) = "M3"
                                       M3Count = M3Count + 1
                                   ElseIf MArr(M) = 2 Then
                                       MArr(M) = "M6"
                                       M6Count = M6Count + 1
                                   Else
                                       haveNXT = Fale


                                       If MArr(M) = "B" Then
                                           BBCount = BBCount + 1
                                       ElseIf MArr(M) = "X" Then
                                           XCount = XCount + 1
                                       End If


                                   End If
                               Next M



                               If haveNXT = True Then
                                   MEndType = Join(MArr, " ")
                                   'MsgBox (MendType)
                                   MEndType = MEndType & Chr(10) & Trim(Str(M3Count)) & "个M3" & Str(M6Count) & "个M6"
                                   'MEndType = MEndType & Chr(10) & M3Count & "个M3模组" & M6Count & "个M6模组"
                                   Cells(1 + MK, 10 + Mj + 1) = MEndType
                                   Cells(1 + MK, 10 + Mj + 1 + 1) = M3Count + M6Count
                               Else
                                   MEndType = Join(MArr, "")
                                   Cells(1 + MK, 10 + Mj + 1) = MEndType
                                   If MArr(0) = "B" Then
                                       Cells(1 + MK, 10 + Mj + 1 + 1) = BBCount / 2
                                   Else
                                       Cells(1 + MK, 10 + Mj + 1 + 1) = 1
                                   End If


                               End If

                           '水平居左
                           'Cells(1 + MK, 10 + Mj + 1 + 1).HorizontalAlignment = xlLeft


                           MK = MK + 1 '先横向列加1，
                           If max<MK Then
                               max = MK
                           End If
                       End If
                       'M行数据复制到列


                       If Cells(i, 1) = "T" Then

                           'Tj = 6 + 1
                           Tj = Mj + UBound(Mtitle)

                           'TK = anotherLstart

                               Cells(1 + TK, 10 + Tj) = Cells(i, 2)
                               Cells(1 + TK, 10 + Tj + 1) = Cells(i, 3)



                           TK = TK + 1
                           If max<TK Then
                               max = TK
                           End If
                       End If
                       'T行数据复制到列






                       If Cells(i, 1) = "B" Then

                           'Bj = 8 + 1
                           Bj = Tj + UBound(Ttitle)
                           'BK = anotherLstart

                               Cells(1 + BK, 10 + Bj) = Cells(i, 2)
                               Cells(1 + BK, 10 + Bj + 1) = Cells(i, 3)
                               Cells(1 + BK, 10 + Bj + 2) = Cells(i, 5)
                               Cells(1 + BK, 10 + Bj + 3) = Cells(i, 6)
                               Cells(1 + BK, 10 + Bj + 4) = Cells(i, 7)


                           BK = BK + 1
                           If max<BK Then
                               max = BK
                           End If
                       End If
                       'M行数据复制到列




                       i = i + 1
                       Wend
                       ' L行开始的一组循环结束


                   condit = i
                   '更新L行的一组驯韩结束后所到达的行号



               Wend

               'Dim Ltitle(0 To 2) As String, Mtitle(3), Ttitle(2), Btitle(5)
               '调用函数可以使用call也可以不使用call，但是最后调用函数和子过程都使用call，否则会提示缺少=
               Call firstTitle(10, 2, Ltitle)
               Call firstTitle(10, 2 + UBound(Ltitle), Mtitle)
               Call firstTitle(10, 2 + UBound(Ltitle) + UBound(Mtitle), Ttitle)
               Call firstTitle(10, 2 + UBound(Ltitle) + UBound(Mtitle) + UBound(Ttitle), Btitle)
               */
    }
}
