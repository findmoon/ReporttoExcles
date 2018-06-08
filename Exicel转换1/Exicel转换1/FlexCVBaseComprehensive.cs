using Exicel转换1;
using NPOIUse;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    //综合的summaryinfo和layoutinfo的类，获取的是完全计算好的两个表和图片、线体信息
    class FlexCVBaseComprehensive
    {
        private SummaryInfoList summaryInfoList_File;
        private layoutInfoList layoutInfoList_File;
        public List<BaseComprehensive> baseComprehensive_list = new List<BaseComprehensive>();
        /*
        //存放SummaryInfoList中的SummaryEveryInfo_list，因为可能会打开多个SummaryInfoList，所以要lsit类型
        List<SummaryEveryInfo> SummaryEveryInfo_list = new List<SummaryEveryInfo>();
        //存放layoutinfolist中的layoutExpressEveryInfo_list
        List<layoutExpressEveryInfo> layoutExpressEveryInfo_list = new List<layoutExpressEveryInfo>();
        */

        //使用元组tuple和list列表存放SummaryEveryInfo，layoutExpressEveryInfo中一一对应的关系
        //List<Tuple<SummaryEveryInfo, layoutExpressEveryInfo>> resultInfoTable_tuple_list = new List<Tuple<SummaryEveryInfo, layoutExpressEveryInfo>>();
        //使用新的类，或者construct，继承 SummaryEveryInfo layoutExpressEveryInfo


        public FlexCVBaseComprehensive(ExcelHelper excelHelper)
        {
            //const string summaryInfoList_File = "summaryInfoList_File" + hadOpenFileNum;
            summaryInfoList_File = new SummaryInfoList(excelHelper);
            summaryInfoList_File.getEveryInfo();

            /*
            for (int i = 0; i < summaryInfoList_File.SummaryEveryInfo_list.Count; i++)
            {
                SummaryEveryInfo_list.Add(summaryInfoList_File.SummaryEveryInfo_list[i]);
            }
            //去掉重复的SummaryEveryInfo，有jobname和target确定唯一的SummaryEveryInfo
            */

            layoutInfoList_File = new layoutInfoList(excelHelper);
            layoutInfoList_File.getEveryInfo();
            /*
            for (int i = 0; i < layoutInfoList_File.layoutExpressEveryInfo_list.Count; i++)
            {
                layoutExpressEveryInfo_list.Add(layoutInfoList_File.layoutExpressEveryInfo_list[i]);
            }
            */
            //在同一个Excel中，Result中job列表的顺序和后面CustomerReport详细信息是一致的
            //即第一个job的SummaryEveryInfo对应第一个layoutExpressEveryInfo，使用tuple将其组合起来
            

        }

        public void getBaseComprehensiveList()
        {
            if (layoutInfoList_File.layoutExpressEveryInfo_list.Count == summaryInfoList_File.expressSummaryEveryInfo_list.Count)
            {
                int addLenght = layoutInfoList_File.layoutExpressEveryInfo_list.Count;
                for (int i = 0; i < addLenght; i++)
                {
                    //resultInfoTable_tuple_list.Add(new Tuple<SummaryEveryInfo, layoutExpressEveryInfo>(
                    // summaryInfoList_File.SummaryEveryInfo_list[i], layoutInfoList_File.layoutExpressEveryInfo_list[i]));
                    BaseComprehensive baseComprehensive = new BaseComprehensive(ComprehensiveStaticClass.GetTimeStamp());
                    baseComprehensive.Jobname = summaryInfoList_File.expressSummaryEveryInfo_list[i].Jobname;
                    baseComprehensive.T_or_B = summaryInfoList_File.expressSummaryEveryInfo_list[i].T_or_B;
                    baseComprehensive.TargetConveyor = summaryInfoList_File.expressSummaryEveryInfo_list[i].TargetConveyor;
                    baseComprehensive.Conveyor = summaryInfoList_File.expressSummaryEveryInfo_list[i].Conveyor;
                    baseComprehensive.ModuleCount = summaryInfoList_File.expressSummaryEveryInfo_list[i].ModuleCount;
                    baseComprehensive.SummaryInfoDT = summaryInfoList_File.expressSummaryEveryInfo_list[i].SummaryInfoDT;
                    baseComprehensive.PictureDataByte = summaryInfoList_File.expressSummaryEveryInfo_list[i].PictureDataByte;
                    
                    //baseComprehensive.layoutDT = layoutInfoList_File.layoutExpressEveryInfo_list[i].layoutExpressDT;

                    baseComprehensive.feederNozzleDT = layoutInfoList_File.layoutExpressEveryInfo_list[i].feederNozzleDT;
                    baseComprehensive.totalfeederNozzleDT = layoutInfoList_File.layoutExpressEveryInfo_list[i].totalfeederNozzleDT;

                    baseComprehensive_list.Add(baseComprehensive);
                }
            }
            else
            {
                MessageBox.Show("Excel报告job数量和CustomerReport存在差异。/n请确认Excel表格后重新选择！");

                //执行析构函数
            }

            
        }


        
    }
    
}
