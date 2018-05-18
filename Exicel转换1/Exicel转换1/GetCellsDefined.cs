using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    /// <summary>
    /// 获取指定的单元格(区域)的内容
    /// 只做一件事，获取从指定位置开始，指定位置结束（或指定长宽的DataTable区域）
    /// 1、从指定的DataTable中，获取指定位置的一个单元格区域
    /// 2、获取startPoint位置开始，指定长宽的区域，同时指定是否将第一行作为列
    /// 3、获取startPoint位置开始，endPoint位置结束的中间区域，同时指定是否将第一行作为列
    /// 4、若开始和结束位置处于同一行，指定是否查找此行下面，两点中间所有的区域，即两列之间区域
    ///    若处于同一列，指定是否查找两点右边，位于两点之间所有行的区域
    /// </summary>
    //
    class GetCellsDefined
    {
        private DataTable _dt;

        //只读属性，仅仅获取指定后的区域，不做其他操作。
        public DataTable DT
        {
            get
            {
                return this._dt;
            }
        }

        #region 从dt中获取，从startPoint位置开始，开始、结束位置的区域
        //从dt中获取，从startPoint位置开始，rowY行、columnX列的单元格内容
        //isFirstRowColumn 参数，是否将这一区域的第一行作为列名
        public DataTable GetCellArea(DataTable dt, int[] startPoint, int rowY, int columnX)
        {
            return GetCellArea(dt, startPoint, rowY, columnX, false);
        }

        public DataTable GetCellArea(DataTable dt,int[] startPoint,int rowY, int columnX, bool isFirstRowColumn)
        {
            DataTable newDT = new DataTable();
           
            DataRow newDR;

            //开始获取的基准点
            int startX = startPoint[0];
            int startY = startPoint[1];

            ///初始化列
            ///仅初始化一次列，如果第一行作为列名isFirstRowColumn为真，则第一行初始化为列名；否则初始化为空列。
            if (!isFirstRowColumn)
            {
                for (int j = 0; j < columnX; j++)
                {
                    newDT.Columns.Add();
                    /*
                        if (isFirstRowColumn)
                        {
                            
                        }
                        else
                        {
                            newDT.Columns.Add();
                        }
                        */
                }
            }
            else
            {
                for (int j = 0; j < columnX; j++)
                {
                    //newDC = new DataColumn(dt.Rows[0][j].ToString());
                    string newDC = dt.Rows[startX][startY + j].ToString();
                    newDT.Columns.Add(newDC);                    
                }
                startX++;
                rowY--;
            }

                ///往行内填充数据
                ///
            for (int i = 0; i < rowY; i++)
            {
                //创建行newDR
                newDR = newDT.NewRow();
                for (int j = 0; j < columnX; j++)
                {
                    if (!dt.Rows[startX + i][startY + j].Equals(DBNull.Value))
                    {
                        //获取数据到newDR
                        newDR[j] = dt.Rows[startX + i][startY + j];
                        //MessageBox.Show(newDR[j].ToString());
                    }
                    
                }

                //将行添加到DataTable
                newDT.Rows.Add(newDR);
            }

            return newDT;
        }

        //获取指定的开始和结束的位置的区域
        /////判断是否有结束位置
        ///若开始和结束位置处于同一行，默认查找此行下面，两点中间所有的区域，即两列之间区域
        ///若处于同一列，默认查找两点右边，位于两点之间所有行的区域
        ///  
        /// 属于通用方法
        /// 
        public DataTable GetCellArea(DataTable dt, int[] startPoint, int[] endPoint, bool isFirstRowColumn)
        {
            //判断是否有结束位置
            if (endPoint==null)
            {
                return GetCellArea(dt, startPoint, dt.Rows.Count - startPoint[0], dt.Columns.Count - startPoint[1], isFirstRowColumn);
            }
            else
            {
                ///如果有结束位置，判断是否同一行同一列
                //判断是否位于同一行或同一列
                //如果位于同一行，获取两点间所有列
                if (startPoint[0] == endPoint[0])
                {
                   return GetCellArea(dt, startPoint, dt.Rows.Count - startPoint[0], endPoint[1] - startPoint[1], isFirstRowColumn);

                }
                else if(startPoint[1] == endPoint[1])//如果位于同一列，获取两点间所有行
                {
                   return GetCellArea(dt, startPoint, endPoint[0] - startPoint[0], dt.Columns.Count - startPoint[1], isFirstRowColumn);

                }
                else
                {
                    //非同一行或同一列，正常调用
                    return GetCellArea(dt, startPoint, endPoint[0] - startPoint[0], endPoint[1] - startPoint[1], isFirstRowColumn);
                }

            }

            
        }

       
        public DataTable GetCellArea(DataTable dt, int[] startPoint, int[] endPoint)
        {
            
            return GetCellArea(dt, startPoint, endPoint, false);

        }

        #endregion

        //获取dt下，开始位置，结束位置 根据下一行是否为空确定
        public DataTable GetCellArea(DataTable dt, int[] startPoint,int columnNum, bool isFirstRowColumn)
        {
            //计算结束位置
            int i = startPoint[0];
            while (!dt.Rows[i][startPoint[1]].Equals(DBNull.Value))
            {
                i++;
            }

            int[] endPoint = new int[2] {i-1, startPoint[1]+ columnNum };

            return GetCellArea(dt, startPoint, endPoint, isFirstRowColumn);
        }
        public DataTable GetCellArea(DataTable dt, int[] startPoint, int columnNum)
        {
            return GetCellArea(dt, startPoint, columnNum,false);
        }

    }
}
