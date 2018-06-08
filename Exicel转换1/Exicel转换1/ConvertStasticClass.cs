using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Exicel转换1
{
    static class ConvertStasticClass
    {
        //静态方法，获取拼接的line
        public static string GetJointLine(Dictionary<string, int> module_Statistics, string machineKind,
            out int Count_M3,out int Count_M6)
        {
            #region//拼接Line字符串

            //重新计算base信息
            //存储模组对应Base的个数，1个M3 module 对应1个M的base，2个M3 module或1个M6 Module对应2Mbase，4个M3 module或2个M6 Module对应4Mbase
            //base尽可能少的分配
            int Count_M = 0;
            //单独统计M#、M6个数
            Count_M3 = 0;
            Count_M6 = 0;
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
            return Line;
            #endregion
        }
    }
}
