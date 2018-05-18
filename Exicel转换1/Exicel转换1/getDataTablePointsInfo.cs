using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    /// <summary>
    /// 实现的功能：
    /// 获取位置
    /// 1、根据输入的字符串，查找字符串位置
    /// 2、根据输入的位置（坐标x\y），查找里面的内容-字符串
    /// 3、根据输入的区间（开始、结束位置的字符串:开始、结束位置的坐标），查找开始结束位置对应的坐标:字符串
    /// 4、给定区间，获取这个区间内的某个字符串的位置，（区间分为位置区间，内容区间,实例化后调用方法一样）
    /// 5、给定区间，获取这个区间内的某个位置的字符串《暂时不实现》
    /// </summary>
    //查找Excel中某列或某行区间之间信息的类,
    //根据位置获取string，或根据string获取位置
    class getDataTablePointsInfo
    {
        #region 设置属性和字段
        //要查找开始结束区间的字符串
        private string _startString;
        private string _endString;

        //开始结束区间对应的位置
        private int[] _startPoint=null;
        private int[] _endPoint=null;

        //查找的DataTable
        private DataTable _dt;

        public DataTable DT
        {
            set
            {
                _dt = value;
            }
            get
            {
                return _dt;
            }
        }

        //设置区间字符串和位置的属性
        public string StartString
        {
            set
            {
                _startString = value;
            }

            get
            {
                return _startString;
            }
        }
        public string EndString
        {
            set
            {
                _endString = value;
            }

            get
            {
                return _endString;
            }
        }

        public int[] StartPoint
        {
            set
            {
                _startPoint = value;
            }

            get
            {
                return _startPoint;
            }
        }
        public int[] EndPoint
        {
            set
            {
                _endPoint = value;
            }

            get
            {
                return _endPoint;
            }
        }
        #endregion

        //构造函数，根据查找的字符串获取位置
        //获取某个字符串在dt中的位置
        public getDataTablePointsInfo(DataTable dt, string startString):this(dt, startString, null)
        {
            

        }

        //获取开始和结束字符串的位置
        public getDataTablePointsInfo(DataTable dt,string startString,string endString)
        {
            //判断dt是否为空
            if (dt!=null)
            {
                this.DT = dt;

                //DataTable中数据索引也是从0，0开始，所以初始化为-1，在获取时保证获取到实际的索引位置
                //pointX\pointY记录在dt中的位置
                int pointX =0;

                //标记获取的开始和结束位置是否唯一
                bool isStartUnique = true;
                bool isEndUnique = true;

                foreach (DataRow dr in dt.Rows)
                {
                    
                    int pointY = 0;
                    foreach (var cellData in dr.ItemArray)
                    {
                        
                        if (cellData.ToString()==startString)
                        {
                            //返回开始位置
                            //当有多个开始位置时会覆盖值，所以需要保证开始和结束位置在Excel唯一
                            if (isStartUnique)
                            {
                                this.StartPoint = new int[] { pointX, pointY };
                                this.StartString = startString;
                                isStartUnique = false;
                            }
                            else
                            {
                                MessageBox.Show("要查找的开始位置或内容必须唯一", "Error!", MessageBoxButtons.OKCancel);

                                //取消初始化，需要特殊处理下，是否释放资源
                                this.StartPoint = null;
                                this.StartString = null;
                                return;
                            }
                            
                        }

                        if (endString!=null)
                        {
                            if (cellData.ToString() == endString)
                            {
                                //返回结束位置
                                if (isEndUnique)
                                {
                                    this.EndPoint = new int[] { pointX, pointY };
                                    this.EndString = endString;
                                    isEndUnique = false;
                                }
                                else
                                {
                                    MessageBox.Show("要查找的结束位置或内容必须唯一", "Error!", MessageBoxButtons.OKCancel);
                                    //取消初始化，需要特殊处理下，是否释放资源
                                    this.EndPoint = null;
                                    this.EndString = null;
                                    return;
                                }
                                
                            }
                        }                       

                        pointY++;
                    }

                    pointX++;
                }

            }
            
        }

        /// <summary>
        /// 构造函数，获取整个DataTable，查找位置也是整个DataTable
        /// </summary>
        public getDataTablePointsInfo(DataTable dt)
        {
            this.DT = dt;
            //要查找的开始和结束位置shi0,0,和最后，及其对应的（数据）字符串
            this.StartPoint =new int[]{ 0,0};
            this.EndPoint = new int[] { dt.Rows.Count-1,dt.Columns.Count-1};
            this.StartString = dt.Rows[0][0].ToString();
            this.EndString = dt.Rows[this.EndPoint[0]][this.EndPoint[1]].ToString();

        }

        ////构造函数，根据位置获取字符串
        //获取某个个字符串在dt中的位置
        public getDataTablePointsInfo(DataTable dt, int[] startPoint) : this(dt, startPoint, null)
        {


        }
        //获取两个字符串 的位置
        public getDataTablePointsInfo(DataTable dt,int[] startPoint,int[] endPoint)
        {
            this.DT = dt;
            //返回要查找的开始和结束位置对应的（数据）字符串
            this.StartString = dt.Rows[startPoint[0]][startPoint[1]].ToString();
            this.StartPoint = startPoint;

            if (endPoint!=null)
            {
                this.EndString = dt.Rows[endPoint[0]][endPoint[1]].ToString();
                this.EndPoint = endPoint;
            }
            
        }

        //获取某个位置间内的字符串对应的位置
        //如果位置内有多个字符串，则应该返回每个字符串对应的位置，即多个位置
        //如果有则返回一个二维数组，没有的返回一个空的二维数组
        //
        //未找到则返回一个length为0 的二维数组
        //isFullFind 是否模糊查找，还是完全相等查找
        public int[][] GetInfo_Between(string findString,bool isFullFind=true)
        {
            #region 切记不要依赖于其他文件、类、或者类库，，不用获取 指定区域 也能实现，一个类只实现一个方法，且要独立
            /*
            //获取实例初始化时指定的开始结束位置间 指定的区域，实现：查找初始化实例后范围内的字符串或信息，不查找其他范围内的数据
            GetCellsDefined getCells = new GetCellsDefined();
            //如果没有结束位置
            getCells.GetCellArea(this.DT, this.StartPoint, this.EndPoint);

            //设定存放多个数组（多个point）的列表
            List<int[]> pointsList = new List<int[]>();

            //获取要查找的区域,和查找区域的大小
            DataTable findDT = getCells.DT;

            int rowCount = findDT.Rows.Count;
            int colCount = findDT.Columns.Count;
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    if (findDT.Rows[i][j].ToString().Contains(findString))
                    {
                        int[] PointXY = new int[2] { this.StartPoint[0] + i, this.StartPoint[1] + j };
                        pointsList.Add(PointXY);
                    }
                }
            }

            return pointsList.ToArray();
            */
            #endregion

            //在类实例化后的对象内（实例化时指定的DataTable内）获取要查找的 字符串或信息，不查找其他范围内的数据

            //如果没有结束位置this.EndPoint=null
            if (this.EndPoint==null)
            {
                this.EndPoint = new int[] { this._dt.Rows.Count-1,this._dt.Columns.Count-1};
            }
            if (this.StartPoint[0]==this.EndPoint[0])//同一行
            {
                this.EndPoint = new int[] { this._dt.Rows.Count-1, this.EndPoint[1] };
            }
            if (this.StartPoint[1] == this.EndPoint[1])//同一列
            {
                this.EndPoint = new int[] { this.EndPoint[0], this._dt.Columns.Count-1 };
            }

            //获取的位置，或多个位置
            List<int[]> pointsList = new List<int[]>();
         

            int findRow = this.EndPoint[0];
            int findCol = this.EndPoint[1];
            for (int i = this.StartPoint[0]; i <= findRow; i++)
            {
                for (int j = this.StartPoint[1]; j <= findCol; j++)
                {
                    if (isFullFind)
                    {
                        if (this._dt.Rows[i][j].ToString()==findString)
                        {
                            int[] PointXY = new int[2] { i, j };
                            pointsList.Add(PointXY);
                        }
                    }
                    else
                    {
                        if (this._dt.Rows[i][j].ToString().Contains(findString))
                        {
                            int[] PointXY = new int[2] { i, j };
                            pointsList.Add(PointXY);
                        }
                    }
                }
            }

            if (pointsList.Count==0)
            {
                MessageBox.Show(string.Format("未找到{0}",findString));
                return null;
            }
            return pointsList.ToArray();
        }


        //获取某个位置间内的位置对应的字符串，位置points是二维数组，唯一

        public string GetInfo_Between(int[] findPoints)
        {
            //位置肯定唯一，所以直接获取即可；
            //判断唯一行是否为DBNull
            if (this.DT.Rows[findPoints[0]][findPoints[1]].Equals(DBNull.Value))
            {
                return null;
            }
            return  this.DT.Rows[findPoints[0]][findPoints[1]].ToString();
            
        }


    }
}
