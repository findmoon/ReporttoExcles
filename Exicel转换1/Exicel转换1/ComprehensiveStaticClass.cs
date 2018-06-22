using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOIUse;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Exicel转换1
{
    //枚举类型，代表横向和竖向的枚举类型，用于插入文字时合并单元格方位,或插入图片时隐藏border边框的选择
    public enum Orientation
    {
        Horizontal,
        Vertical,
        Origin//圆点，表示不合并不横向也不竖向
    }

    //
    //存放module head 理论CPH素有可能性，用以修改的结构
    public struct Module_Head_Cph_Struct
    {
        public string[] moduleTrayTypestring;

        public List<string> headTypeString_List;
        public List<string[]> CPH_strings_List;
    }

    //存放一些公共的静态方法、数据（字典）的静态类
    static class ComprehensiveStaticClass
    {
        
        //电、气每种机器的值是固定的，使用字典，将字典放入结构中
        //所有Tray：LTray-LTTray-LT2Tray-LTCTray-MTray。只在M6II\M6III上搭配Tray
        //功率字典
        static public Dictionary<string, double> power_Dictionary = new Dictionary<string, double>()
        {
            {"M3",0.6 },
            {"M3S",0.8 },
            {"M6",0.6 },
            {"M6S",1.0 },
            {"M3II",1.5 },
            {"M6II",1.5 },
            {"M3IIc",1.5 },
            {"M6IIc",1.5 },
            {"M3III",0.8 },
            {"M6III",0.9 },
            {"M3IIIc",0.8 },
            {"M6IIIc",0.9 },
            {"2MBaseII", 1.0 },
            {"4MBaseII",1.0 },
            {"2MBase", 0.5 },
            {"4MBase",0.7 },//实际上为III代，应该写成全局配置参数，方便调整修改
            {"LTray",0.1 },
            {"LTTray",0.5 },
            {"LT2Tray",0.5 },
            {"LTCTray",0.5 }//M Tray
        };
        static public Dictionary<string, int> air_Consumption_Dictionary = new Dictionary<string, int>()
        {
            {"2MBase", 45},
            {"4MBase",90 }
        };
        //int ip;//一个base一个ip
        static public Dictionary<string, double> baseLength_Dictionary = new Dictionary<string, double>()
        {
            {"2MBase",650},
            {"4MBase",1300}
        };


        #region //根据各种信息生成summary信息的  除cycletime以外的  DataTable
        //无cycletime，board output/hour、board output/day、 cPH\CPH rate,cycletime\cPH需要由layoutinfo 获取
        public static DataTable getExpressSummayDataTable(int boardQty, string Line, string Head_Type,
            string jobName, string pannelSize, string placementNumber, string remark)
        {
            //生成评估的表格
            DataTable genDT = new DataTable();
            string[] genColumnString = { "Item", "Board", "Line", "Head \nType", "Job name","Panel Size",
                "Cycle time\n[((L1+L2)/2) Sec]", "Number of \nplacements", "CPH", "Output \nBoard/hour",
                "Output\nBoard/Day(22H)", "CPH Rate","CPP","Remark" };

            //评估表格的栏
            for (int i = 0; i < genColumnString.Length; i++)
            {
                genDT.Columns.Add(genColumnString[i]);
            }

            //向生成的Excel添加行
            DataRow genRow = genDT.NewRow();
            genRow["Board"] = boardQty;
            genRow["Line"] = Line;
            genRow["Head \nType"] = Head_Type;
            genRow["Job name"] = jobName;
            genRow["Panel Size"] = pannelSize;
            //genRow[5] = cycleTime;
            genRow["Number of \nplacements"] = placementNumber;
            //genRow["CPH"] = cPH;
            //float.Parse(cycleTime)  float转换
            //double doubleOutput = Convert.ToDouble(boardQty) * 3600 / Convert.ToDouble(cycleTime);
            //genRow["Output \nBoard/hour"] = string.Format("{0:0}", doubleOutput);
            // genRow["Output\nBoard/Day(22H)"] = string.Format("{0:0}", doubleOutput * 22);
            //42000  H24头理论的CPH，10500  H04SF理论的CPH
            //需要建一个head type:CPH的字典，进行计算；

            //CPH rate的计算，head的理论CPH*head num ++
            //double CPHRate = Convert.ToDouble(cPH) / (42000 * 7 + 10500 * 1);
            //genRow["CPH Rate"] = string.Format("{0:0.00%}", CPHRate);
            //genRow[11] = "已删减大零件";
            genRow["Remark"] = remark;

            //添加行
            genDT.Rows.Add(genRow);
            return genDT;
        }
        #endregion

        #region 生成summaryinfo的图片
        /// <summary>
        /// 生成图片
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint"></param>
        public static byte[] GenModulesPicture(string[] allModuleTypeString, Dictionary<string, int> module_Statistics)
        {

            string img_M3 = @".\img\M3.png";
            string img_M6 = @".\img\M6.png";
            string img_M6LTray = @".\img\M6-LTray.png";
            string img_M6LTTray = @".\img\M6-LTTray.png";
            string img_M6LT2Tray = @".\img\M6-LT2Tray.png";
            string img_M6LTCTray = @".\img\M6-LTCTray.png";
            string img_M6MTray = @".\img\M6-MTray.png";

            if (!File.Exists(img_M3) || !File.Exists(img_M3)|| !File.Exists(img_M6LTray) || !File.Exists(img_M6LTTray)
                || !File.Exists(img_M6LTCTray) || !File.Exists(img_M6MTray)|| !File.Exists(img_M6LT2Tray))
            {
                MessageBox.Show("缺少模组相关图片，请确认后重试");
                return null;
            }

            //获取 图片的img
            var imgM3 = Image.FromFile(img_M3);
            var imgM6 = Image.FromFile(img_M6);
            //img_M6LTray、img_M6LTTray 、img_M6LTCTray、img_M6MTray
            var imgM6LTray = Image.FromFile(img_M6LTray);
            var imgM6LTTray = Image.FromFile(img_M6LTTray);
            var imgM6LT2Tray = Image.FromFile(img_M6LT2Tray);
            var imgM6LTCTray = Image.FromFile(img_M6LTCTray);
            var imgM6MTray = Image.FromFile(img_M6MTray);

            #region//获取要生成的图片的宽高
            //生成最后图片的外边框留白,jiange间隔
            int jiange = 5;
            //int finalWidth = imgM3.Width * module_Statistics[moduleKind3_III]/2 + imgM6.Width*module_Statistics[moduleKind6_III] + jiange*2;

            //统计所需要拼接的图片数量
            int M3Image_count = 0;
            int M6Image_count = 0;
            //模组M3或M6的个数
            int M3_count=0, M6_count=0;

            foreach (var item in module_Statistics)
            {
                //区别keyM3III\M3II\M3I
                if (item.Key.Substring(0, 2) == "M3")
                {
                    M3_count += item.Value;
                }
                if (item.Key.Substring(0, 2) == "M6")
                {
                    M6_count += item.Value;
                }
            }
            M3Image_count = M3_count / 2;
            M6Image_count = M6_count;
            int finalWidth = imgM3.Width * M3Image_count + imgM6.Width * M6Image_count + jiange * 2;
            int finalHeight = imgM3.Height + jiange * 2;
            #endregion 

           

            Bitmap finalImg = new Bitmap(finalWidth, finalHeight);
            Graphics graph = Graphics.FromImage(finalImg);
            //graph.Clear(SystemColors.AppWorkspace);
            graph.Clear(Color.Empty);

            #region//循环模组并画取img画图
            //获取的位置
            int pointX = jiange;
            int pointY = jiange;

            ////确定最后的模组和Tray,同时画取图像
            for (int i = 0; i < allModuleTypeString.Length; i++)
            {
                //M3模组图像
                if (allModuleTypeString[i].Substring(0, 2) == "M3" && allModuleTypeString[i + 1].Substring(0, 2) == "M3")
                {
                    //写入文字M3或M3S
                    graph.DrawImage(WriteTextToM3Image(allModuleTypeString[i], allModuleTypeString[i+1],imgM3), pointX, pointY);
                    pointX += imgM3.Width;
                    i++;
                }
                else if (allModuleTypeString[i].Substring(0, 2) == "M6")//M6模组图像
                {
                    string[] moduleTypeSplit = allModuleTypeString[i].Split('-');
                    if (moduleTypeSplit.Length>1)//imgM6LTray\imgM6LTTray\imgM6LTCTray\imgM6MTray
                    {
                        switch (moduleTypeSplit[1].ToLower())
                        {
                            case "ltray":
                                {
                                    graph.DrawImage(WriteTextToM6Image(moduleTypeSplit[0],imgM6LTray), pointX, pointY);
                                    pointX += imgM6.Width;
                                    break;
                                }
                            case "mtray":
                                graph.DrawImage(WriteTextToM6Image(moduleTypeSplit[0], imgM6MTray), pointX, pointY);
                                pointX += imgM6.Width;
                                break;
                            case "lttray":
                                graph.DrawImage(WriteTextToM6Image(moduleTypeSplit[0], imgM6LTTray), pointX, pointY);
                                pointX += imgM6.Width;
                                break;
                            case "lt2tray":
                                graph.DrawImage(WriteTextToM6Image(moduleTypeSplit[0], imgM6LT2Tray), pointX, pointY);
                                pointX += imgM6.Width;
                                break;
                            case "ltctray":
                                graph.DrawImage(WriteTextToM6Image(moduleTypeSplit[0], imgM6LTCTray), pointX, pointY);
                                pointX += imgM6.Width;
                                break;
                            default:
                                break;
                        }
                        
                    }
                    else
                    {
                        graph.DrawImage(WriteTextToM6Image(allModuleTypeString[i], imgM6), pointX, pointY);
                        pointX += imgM6.Width;
                    }
                }
            }
            graph.Dispose();
            #endregion

            #region//将bitmap转换为byte[],并返回
            byte[] PictureDataByte;
            using (MemoryStream stream = new MemoryStream())
            {
                finalImg.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                PictureDataByte = new byte[stream.Length];
                stream.Seek(0, SeekOrigin.Begin);
                stream.Read(PictureDataByte, 0, Convert.ToInt32(stream.Length));
            }
            return PictureDataByte;
            #endregion
        }
        #endregion

        private static Image WriteTextToM3Image(string str1, string str2,Image m3Image)
        {
            Font font = new Font("Arial", 8, FontStyle.Regular);
            //绘笔颜色
            SolidBrush brush = new SolidBrush(Color.Black);
            //图画
            Bitmap bitmap = new Bitmap(m3Image, m3Image.Width, m3Image.Height);
            Graphics g = Graphics.FromImage(bitmap);
            //g.Clear(Color.Empty);
            //g.DrawImage(m3Image, 0, 0, m3Image.Width, m3Image.Height);
            //g.Clear(Color.Empty);
            ////定义一个矩形区域，在这个矩形上画字
            float rectX = 10;
            float rectY = 37;
            //字体格式
            //StringFormat format = new StringFormat(StringFormatFlags.NoClip);
            SizeF sizeF = g.MeasureString(str1, font);
            int width = (int)(sizeF.Width+1);
            int height = (int)(sizeF.Height+1);
            //声明矩形区域
            RectangleF textArea = new RectangleF(rectX, rectY, width, height);
            g.DrawString(str1, font, brush,textArea);

            rectX = 50;
            rectY = 37;
            //字体格式
            //StringFormat format = new StringFormat(StringFormatFlags.NoClip);
            sizeF = g.MeasureString(str2, font);
            width = (int)(sizeF.Width+1);
            height = (int)(sizeF.Height+1);
            //声明矩形区域
            textArea = new RectangleF(rectX, rectY, width, height);
            g.DrawString(str2, font, brush,textArea);
            g.Dispose();
            return bitmap;
        }
        private static Image WriteTextToM6Image(string str, Image m6Image)
        {
            Font font = new Font("Arial", 10, FontStyle.Regular);
            //绘笔颜色
            SolidBrush brush = new SolidBrush(Color.Black);
            //图画
            
            Bitmap bitmap = new Bitmap(m6Image, m6Image.Width, m6Image.Height);
            Graphics g = Graphics.FromImage(bitmap);
            //g.Clear(Color.Empty);
            //g.DrawImage(m6Image, 0, 0, m6Image.Width, m6Image.Height);
            float rectX = 15;
            float rectY = 37;
            //字体格式
            //StringFormat format = new StringFormat(StringFormatFlags.NoClip);
            SizeF sizeF = g.MeasureString(str, font);
            int width = (int)(sizeF.Width+1);
            int height = (int)(sizeF.Height+1);
            //声明矩形区域
            RectangleF textArea = new RectangleF(rectX, rectY, width, height);
            g.DrawString(str, font, brush, textArea);
            g.Dispose();
            return bitmap;
        }

        #region 保存对话框的实现
        public static string SaveExcleDialogShow()
        {
            SaveFileDialog saveExcelFileDialog = new SaveFileDialog();
            //保存的路径
            //string fileName = @"D:\Test\110.xlsx";
            string fileName = null;
            saveExcelFileDialog.Filter = "Excel 2007 工作表(*.xlsx)|*.xlsx|Excel 97-2003 工作表(*.xls)|*.xls";
            saveExcelFileDialog.RestoreDirectory = true;
            if (saveExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    fileName = saveExcelFileDialog.FileName;
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }
                    return fileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return null;
                }

            }
            else
            {
                return null; ;
            }
        }
        #endregion
        
        #region //将string插入Excel指定位置
        private static void StringInsertWorkbook(ExcelHelper excelHelper1, string insertString, string sheetName,
            int[] insertPoint)
        {
            StringInsertWorkbook(excelHelper1, insertString, sheetName, insertPoint, Orientation.Origin, 0, null, 0);
        }
        #endregion
        #region //将string插入Excel指定位置
        /// <summary>
        /// 将字符创插入指定sheetname表单中指定的单元格
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="insertString"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint">插入位置</param>
        /// <param name="hasBorder">是否设置边框，默认不设置</param>
        private static void StringInsertWorkbook(ExcelHelper excelHelper1, string insertString, string sheetName,
            int[] insertPoint, bool hasBorder=false)
        {
            StringInsertWorkbook(excelHelper1, insertString, sheetName, insertPoint, Orientation.Origin, 0, null,0, hasBorder);
        }

        //重载字符串插入sheet，指定合并插入时的单元格，合并方向和单元格数量
        //MergeOrientation
        /// <summary>
        /// 将字符创插入指定sheetname表单中指定的单元格，并合并指定单元格
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="insertString"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint"></param>        
        /// <param name="mergeOrientation"></param>
        /// <param name="mergeNum">从插入位置开始合并单元格的数量</param>
        /// <param name="WidthHeight">插入string所在单元格的宽高，只实现了设置高度,null表示不设置宽高</param>
        /// <param name="colorShort">colorShort NPOI颜色，0表示不设置颜色</param>
        /// <param name="hasBorder">是否显示边框，默认不显示</param>
        private static void StringInsertWorkbook(ExcelHelper excelHelper1, string insertString, string sheetName,
            int[] insertPoint, Orientation mergeOrientation,int mergeNum, int[] WidthHeight, short colorShort,bool hasBorder=true)
        {
            IWorkbook workbook = excelHelper1.WorkBook;
            //FileStream fs = excelHelper1.FS;

            //从指定的位置开始插入string
            int count = insertPoint[0];
            ISheet sheet = excelHelper1.ExcelToIsheet(sheetName);

            //sheet由excelHelper获取
            try
            {
                IRow row;
                if (sheet.GetRow(count) == null)
                {
                    row = sheet.CreateRow(count);
                }
                else
                {
                    row = sheet.GetRow(count);
                }
                if (WidthHeight!=null)
                {
                    row.HeightInPoints = WidthHeight[1];
                }                
                row.CreateCell(insertPoint[1]).SetCellValue(insertString);

                //
                #region 样式
                ICellStyle cellStyle = workbook.CreateCellStyle();
                //垂直居中
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //水平对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                cellStyle.WrapText = true;

                //是否设置边框
                if (hasBorder)
                {
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                }

                if (colorShort != 0)
                {
                    cellStyle.FillBackgroundColor = IndexedColors.White.Index;
                    cellStyle.FillForegroundColor = colorShort;
                }

                cellStyle.FillPattern = FillPattern.SolidForeground;
                #endregion
                //合并单元格
                //mergeNum为合并的cell数量，结束位置为mergeNum-1
                if (mergeOrientation==Orientation.Horizontal)//横向
                {                    
                    //设置样式到单元格
                    for (int i = insertPoint[1]; i < insertPoint[1]+ mergeNum; i++)
                    {
                        if (row.GetCell(i)==null)
                        {
                            row.CreateCell(i);
                        }
                        row.GetCell(i).CellStyle = cellStyle;
                        
                    }
                    //合并
                    sheet.AddMergedRegion(new CellRangeAddress(insertPoint[0], insertPoint[0], insertPoint[1], insertPoint[1] + mergeNum-1));
                }
                if (mergeOrientation == Orientation.Vertical)
                {
                    for (int i = insertPoint[0]; i < insertPoint[0]+mergeNum; i++)
                    {
                        if (sheet.GetRow(i)==null)
                        {
                            sheet.CreateRow(i);
                        }
                        if (sheet.GetRow(i).GetCell(insertPoint[1])==null)
                        {
                            sheet.GetRow(i).CreateCell(insertPoint[1]);
                        }
                        sheet.GetRow(i).GetCell(insertPoint[1]).CellStyle = cellStyle;
                    }
                    sheet.AddMergedRegion(new CellRangeAddress(insertPoint[0], insertPoint[0] + mergeNum-1, insertPoint[1], insertPoint[1]));
                }

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        /// <summary>
        /// 生成sheet1 的DataTable
        /// </summary>
        /// <param name="excelHelper1">要生成的Excel对应的ExcelHelper类对象</param>
        /// <param name="dt">生成sheet的DataTable</param>
        /// ///////////////////<param name="fileName">生成的文件，在判断fs没有的情况下，创建fs文件流</param>
        /// <param name="sheetNamet">将dt插入到指定sheet</param>
        /// <param name="isColumnWritten">是否吧列名作为单元格插入</param>
        /// <param name="insertPoint">插入的位置</param>
        /// <param name="isDisplayGrid">插入的位置</param>
        private static void DataTableToResultSheet1(ExcelHelper excelHelper1, DataTable dt, string sheetName, 
            bool isColumnWritten, int[] insertPoint,bool isDisplayGrid=false)
        {
            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;
            //fs 作为向现有文件生成表格还是新建文件生成表格
            //FileStream fs = excelHelper1.FS;

            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];
            ISheet sheet = excelHelper1.ExcelToIsheet(sheetName);

            #region sheet由excelHelper获取
            //if (workbook != null)
            //{
            //    //有则获取
            //    if (workbook.GetSheetIndex(sheetNamet) >= 0)
            //    {
            //        sheet = workbook.GetSheet(sheetNamet);
            //    }
            //    else
            //    {
            //        sheet = workbook.CreateSheet(sheetNamet);
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("失败");
            //    return;
            //}

            #endregion
            try
            {
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Columns[j].ColumnName);
                    }
                    count++;
                }

                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                    ++count;
                }
                //是否显示网格
                if (!isDisplayGrid)
                {
                    sheet.DisplayGridlines = false;
                }
                #region sheet1的样式
                //设置列宽
                int commonWidth = 11;
                sheet.SetColumnWidth(0, 1 * 256);
                sheet.SetColumnWidth(1, commonWidth * 256);
                sheet.SetColumnWidth(2, commonWidth * 256);
                sheet.SetColumnWidth(3, System.Convert.ToInt32(System.Convert.ToDouble(commonWidth * 256) * 3));
                sheet.SetColumnWidth(4, commonWidth * 2 * 256);
                sheet.SetColumnWidth(5, System.Convert.ToInt32(System.Convert.ToDouble(commonWidth  * 256) * 1.8));
                sheet.SetColumnWidth(6, commonWidth * 256);
                sheet.SetColumnWidth(7, commonWidth * 256);
                sheet.SetColumnWidth(8, commonWidth * 256);
                sheet.SetColumnWidth(9, commonWidth * 256);
                sheet.SetColumnWidth(10, commonWidth * 256);
                sheet.SetColumnWidth(11, commonWidth * 256);
                sheet.SetColumnWidth(12, commonWidth * 256);

                //设置行高,仅设置summaryDT两行的行高,仅仅设置 insertPoint[0] 和 count 之间的 行高

                //int rowNum = count - insertPoint[0];
                IRow row1;
                IRow row2;

                if (sheet.GetRow(insertPoint[0]) == null)
                {
                    row1 = sheet.CreateRow(insertPoint[0]);
                }
                else
                {
                    row1 = sheet.GetRow(insertPoint[0]);
                }
                row1.HeightInPoints = 50;
                
                if (isColumnWritten)
                {

                    if (sheet.GetRow(insertPoint[0] + 1) == null)
                    {
                        row2 = sheet.CreateRow(insertPoint[0] + 1);
                    }
                    else
                    {
                        row2 = sheet.GetRow(insertPoint[0] + 1);
                    }
                    row2.HeightInPoints = 50;
                }

                //设置整列单元格样式
                //获取列数量
                //int column_count = ExcelHelper.sheetColumns(sheet);
                //for (int i = 0; i < column_count; i++)
                //{
                //    //设置SetDefaultColumnStyle后，仅到新创建cell时，会应用这个默认设置
                //    //sheet.SetDefaultColumnStyle(i, cellStyle);
                //}
                ICellStyle cellStyle;
                for (int i = insertPoint[0]; i < count; i++)
                {
                    if (sheet.GetRow(i) != null)
                    {
                        IRow row = sheet.GetRow(i);

                        //row.LastCellNum=14，但实际有13个cell,
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //int j = row.FirstCellNum,这个复制在第一次循环时提示cell number必须>=0，FirstCellNum怎么会<0？
                            if (row.GetCell(j) == null)
                            {
                                continue;
                            }

                            //设置单元格样式,必须每一个cell对应一个样式，否则无法设置背景

                            cellStyle = workbook.CreateCellStyle();
                            //垂直居中
                            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                            //水平对齐
                            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                            //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                            cellStyle.WrapText = true;

                            //设置cellstyle背景
                            //cellStyle.FillBackgroundColor = IndexedColors.Automatic.Index;
                            //cellStyle.FillForegroundColor = IndexedColors.Automatic.Index;

                            //设置边框
                            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;


                            if (i == sheet.LastRowNum)
                            {
                                //设置数据含前景色白色
                                //cellStyle.FillForegroundColor = IndexedColors.White.Index;
                                //最后的的remark数据行，设置字体
                                //使用cells.Count可以进入最后一个数据
                                if (j == row.LastCellNum-1)
                                {
                                    //remark左对齐
                                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                                    //创建字体和设置字号
                                    IFont font = workbook.CreateFont();
                                    font.FontHeightInPoints = 8;
                                    cellStyle.SetFont(font);
                                }
                            }

                            //填充模式
                            cellStyle.FillPattern = FillPattern.SolidForeground;

                            row.GetCell(j).CellStyle = cellStyle;
                            //cellStyle.Dispose();
                        }
                    }
                }
                #endregion

                #region 写入流应该在方法外
                //if (fs == null)
                //{
                //    fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //}
                //workbook.Write(fs); //写入到excel文件，并关闭流
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        /// <summary>
        /// 生成sheet2
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="dt"></param>
        /// <param name="sheetNamet"></param>
        /// <param name="isColumnWritten"></param>
        /// <param name="insertPoint"></param>
        /// <param name="isDisplayGrid"></param>
        private static void DataTableToResultSheet2(ExcelHelper excelHelper1, DataTable dt, string sheetNamet, 
            bool isColumnWritten, int[] insertPoint, bool isDisplayGrid = false)
        {
            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;

            //从指定的位置开始插入DataTable数据
            int count = insertPoint[0];

            ISheet sheet = excelHelper1.ExcelToIsheet(sheetNamet);

            try
            {
                IRow row;//声明row 行变量
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    //获取或创建行
                    if (sheet.GetRow(count) == null)
                    {
                        row = sheet.CreateRow(count);
                    }
                    else
                    {
                        row = sheet.GetRow(count);
                    }
                    //将列名插入sheet
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Columns[j].ColumnName);
                    }
                    count++;
                }

                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    //获取或创建行
                    if (sheet.GetRow(count) == null)
                    {
                        row = sheet.CreateRow(count);
                    }
                    else
                    {
                        row = sheet.GetRow(count);
                    }
                    //将内容插入sheet
                    for (int j = 0; j < dt.Columns.Count; ++j)
                    {
                        row.CreateCell(insertPoint[1] + j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                    ++count;
                }

                //setSecondSheetStyle(sheet, dt.Columns.Count - 1);

                if (!isDisplayGrid)
                {
                    sheet.DisplayGridlines = false;
                }

                //设置样式
                //设置行高,仅设置summaryDT两行的行高,仅仅设置 insertPoint[0] 和 count 之间的 行高

                int rowNum = count - insertPoint[0];
                IRow row2;
                if (isColumnWritten)
                {

                    if (sheet.GetRow(insertPoint[0]) == null)
                    {
                        row2 = sheet.CreateRow(insertPoint[0] + 1);
                    }
                    else
                    {
                        row2 = sheet.GetRow(insertPoint[0]);
                    }
                    row2.HeightInPoints = 50;
                }


                //设置整列单元格样式
                //获取列数量
                //int column_count = ExcelHelper.sheetColumns(sheet);
                //for (int i = 0; i < column_count; i++)
                //{
                //    //设置SetDefaultColumnStyle后，仅到新创建cell时，会应用这个默认设置
                //    //sheet.SetDefaultColumnStyle(i, cellStyle);
                //}
                ICellStyle cellStyle;
                for (int i = insertPoint[0]; i < count; i++)
                {
                    if (sheet.GetRow(i) != null)
                    {
                        row = sheet.GetRow(i);

                        //row.LastCellNum=14，但实际有13个cell,
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //int j = row.FirstCellNum,这个复制在第一次循环时提示cell number必须>=0，FirstCellNum怎么会<0？
                            if (row.GetCell(j) == null)
                            {
                                continue;
                            }

                            //设置单元格样式,必须每一个cell对应一个样式，否则无法设置背景

                            cellStyle = workbook.CreateCellStyle();

                            //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                            //第一行换行,且居中
                            if (i== insertPoint[0])
                            {
                                cellStyle.WrapText = true;
                                //垂直居中
                                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                                //水平对齐
                                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            }
                            

                            //设置cellstyle背景
                            //cellStyle.FillBackgroundColor = IndexedColors.Automatic.Index;
                            //cellStyle.FillForegroundColor = IndexedColors.Automatic.Index;

                            //设置边框
                            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;


                            //填充模式
                            cellStyle.FillPattern = FillPattern.SolidForeground;

                            row.GetCell(j).CellStyle = cellStyle;
                            //cellStyle.Dispose();
                        }
                    }
                }

                #region 写入流应该在方法外
                //if (fs == null)
                //{
                //    fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //}
                //workbook.Write(fs); //写入到excel文件，并关闭流
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        #region 设置layout sheet的样式
        private static void SetSecondSheetStyle(ISheet sheet, int mergeLastCount)
        {
            //应使用静态方法，可以获取不规则行列的sheet的最大与最小列
            int sheetColumns = ExcelHelper.sheetColumns(sheet);
            #region sheet2的样式
            //设置列宽
            int commonWidth = 8;
            sheet.SetColumnWidth(0, 1 * 256);
            for (int i = 1; i < sheetColumns; i++)
            {
                if (i == sheetColumns - 2)
                {
                    sheet.SetColumnWidth(i, System.Convert.ToInt32(commonWidth * 1.8) * 256);
                    continue;
                }
                if (i == sheetColumns - 4)
                {
                    sheet.SetColumnWidth(i, System.Convert.ToInt32(commonWidth * 1.8) * 256);
                    continue;
                }
                sheet.SetColumnWidth(i, commonWidth * 256);
            }

            //设置行高
            //是否是有内容的第一行，有内容的第一行上面和左边合并
            bool isFirstDataRow = true;
            //workbook，创建cellstyle
            IWorkbook workbook = sheet.Workbook;
            ICellStyle cellStyle;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                if (sheet.GetRow(i) != null)
                {
                    //设置第一次数据行的行高和和合并，在cell循环之外。
                    if (isFirstDataRow)
                    {
                        sheet.GetRow(i).HeightInPoints = 50;
                        //第一次运行完后，变为非第一行
                        isFirstDataRow = false;
                    }
                    else
                    {
                        sheet.GetRow(i).HeightInPoints = 20;
                    }

                    //设置cell的cellstyle
                    IRow row = sheet.GetRow(i);
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        if (row.GetCell(j) == null)
                        {
                            continue;
                        }
                        //设置单元格样式,必须每一个cell对应一个样式，否则无法设置背景
                        cellStyle = workbook.CreateCellStyle();
                        //垂直居中
                        cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                        //水平对齐
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        //  cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                        //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                        cellStyle.WrapText = true;

                        //设置cellstyle背景
                        cellStyle.FillBackgroundColor = IndexedColors.Black.Index;
                        cellStyle.FillForegroundColor = IndexedColors.White.Index;

                        //设置边框
                        cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                        cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

                        //最后一行设置字体 红色
                        if (i == sheet.LastRowNum)
                        {
                            //创建字体和设置字号
                            IFont font = workbook.CreateFont();
                            font.Color = IndexedColors.Red.Index;
                            cellStyle.SetFont(font);
                        }

                        //填充模式
                        cellStyle.FillPattern = FillPattern.SolidForeground;

                        row.GetCell(j).CellStyle = cellStyle;
                    }
                }

            }
            #endregion

        }
        #endregion


        /// <summary>
        /// 插入图片在ExcelHelper中sheetName表中
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint">插入位置</param>
        /// <param name="PictureDataByte">图片</param>
        /// <param name="machineCount">根据机器长度判定结束单元格的长短</param>
        private static void insertPictureToWorbook(ExcelHelper excelHelper1,string sheetName,int[] insertPoint,
            byte[] PictureDataByte,int machineCount,Orientation hideborderOrientation, int hideborderCellNum, int[] WidthHeight = null)
        {
            #region 将图片插入sheet的实现
            //生成报告的excelHelper1 excelHelper类
            IWorkbook workbook = excelHelper1.WorkBook;

            ISheet sheet1 = excelHelper1.ExcelToIsheet(sheetName);

            //设置行高
            if (WidthHeight!=null)
            {
                IRow row = sheet1.GetRow(insertPoint[0]);
                if (row == null)
                {
                    row = sheet1.CreateRow(insertPoint[0]);
                }
                row.HeightInPoints = WidthHeight[1];
            }

            //从指定的位置开始插入图片
            //int count = insertPoint[0];
            //插入的位置
            int startRow = insertPoint[0];
            //插入图片位置偏移一下，否则样式不好看
            int startCol = insertPoint[1]+ hideborderCellNum/2-2;
            int endRow = startRow + 1;
            int endCol = startCol + 1;
            //根据模组数量长度，变更结束单元格位置

            if (machineCount > 3)
            {
                endCol = startCol + 2;
            }
            if (machineCount > 7)
            {
                endCol = startCol + 3;
            }
            if (machineCount > 10)
            {
                startCol--;
                endCol = startCol + 4;
            }
            if (machineCount > 15)
            {
                startCol--;
                endCol = startCol + 5;
            }

            //偏移依旧不起作用
            excelHelper1.pictureDataToSheet(sheet1, PictureDataByte, 50, 10, 50, 10, startRow, startCol, endRow, endCol);

            //对单元格隐藏的处理
            if (hideborderOrientation==Orientation.Horizontal)
            {
                IRow rowPicture = sheet1.GetRow(insertPoint[0]);
                for (int i = 1; i < hideborderCellNum-1; i++)
                {
                    //设置单元格边框右边为空
                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
                    //填充模式
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    //设置样式
                    ICell cellPicture = rowPicture.GetCell(insertPoint[1] + i);
                    if (rowPicture.GetCell(insertPoint[1] + i)==null)
                    {
                        cellPicture = rowPicture.CreateCell(insertPoint[1] + i);
                    }
                    cellPicture.CellStyle = cellStyle;
                }
            }
            #endregion
        }

        /// <summary>
        /// //比较两个string数组是否相等的
        /// </summary>
        /// <param name="string1"></param>
        /// <param name="string2"></param>
        /// <returns>相等返回true</returns>
        static public bool CompareStringArray(string[] string1,string[] string2)
        {
            if (string1.Length!=string2.Length)
            {
                return false;
            }
            for (int i = 0; i < string1.Length; i++)
            {
                if (string1[i]!=string2[i])
                {
                    return false;
                }
            }
            return true;
        }
        //从BaseComprehensive列表List 按同一机种、同一配置生成Excel
        static public void GenExcelfromBaseComprehensiveList(List<BaseComprehensive> baseComprehensiveList,string excelFileName)
        {
            if (baseComprehensiveList.Count==0)
            {
                return;
            }
            if (baseComprehensiveList.Count == 1)//只选择了一个job报告，直接生成
            {
                GenExclefromBaseComprehensive(excelFileName,baseComprehensiveList[0].Jobname, baseComprehensiveList[0]);                
            }

            if (baseComprehensiveList.Count>1)
            {
                //相同机种的BaseComprehensiveList
                List<BaseComprehensive> theSameModel_BaseComprehensiveList = new List<BaseComprehensive>();
                theSameModel_BaseComprehensiveList.Add(baseComprehensiveList[0]);

                #region //将同一机种、同一配置、和其他 进行分类。使用List下标number区分
                //存放同一机种的number位置 数组的 列表，可能有多个 同一机种组成的数组
                List<int[]> theSameModelArray_List = new List<int[]>();
                //存放同一配置的number位置 数组的 列表，可能有多个 同一配置组成的数组
                List<int[]> theSameConfigArray_List = new List<int[]>();
                //剩余既不是同一配置也不是同一机种
                List<int> theOrther_Number = new List<int>();

                //记录已经查找确认过的位置Number，防止重复查找
                List<int> hasConfirmed_Number = new List<int>();
                for (int i = 0; i < baseComprehensiveList.Count; i++)
                {
                    //已经确认过的跳过，防止重复查找
                    if (hasConfirmed_Number.Contains(i))
                    {
                        continue;
                    }

                    //每次循环当前值i后面的之前，都假设i是多个 同一机种的一个
                    List<int> theSameModel_Number = new List<int>();
                    theSameModel_Number.Add(i);
                    for (int j = i + 1; j < baseComprehensiveList.Count; j++)
                    {
                        if (baseComprehensiveList[i].Jobname == baseComprehensiveList[j].Jobname)
                        {
                            //只能判断i j是同一机种，而不同判断其他还有同一机种
                            theSameModel_Number.Add(j);
                            //添加 已经确认的位置j
                            if (!hasConfirmed_Number.Contains(i))
                            {
                                hasConfirmed_Number.Add(i);
                            }
                            hasConfirmed_Number.Add(j);
                        }
                    }

                    //判断是否已经找到i对用的同一机种
                    if (theSameModel_Number.Count > 1)
                    {
                        //i 有对应的同一机种，将其添加到同一机种数组的列表中
                        theSameModelArray_List.Add(theSameModel_Number.ToArray());
                        //继续下一个 i+1的查询
                        continue;
                    }

                    //i 未找到对应的同一机种

                    //每次循环当前值i后面的之前，都假设i是多个 同一机种的一个
                    List<int> theSameConfig_Number = new List<int>();
                    theSameConfig_Number.Add(i);
                    //判断是否有同一配置
                    for (int j = i + 1; j < baseComprehensiveList.Count; j++)
                    {
                        //module和head string[]相等，判定为同一配置
                        if (CompareStringArray(baseComprehensiveList[i].AllModuleType, baseComprehensiveList[j].AllModuleType) &&
                            CompareStringArray(baseComprehensiveList[i].AllHeadType, baseComprehensiveList[j].AllHeadType))
                        {
                            theSameConfig_Number.Add(j);
                            //添加 已经确认的位置j
                            if (!hasConfirmed_Number.Contains(i))
                            {
                                hasConfirmed_Number.Add(i);
                            }
                            hasConfirmed_Number.Add(j);
                        }
                    }

                    //判定i已找到对应的同一配置
                    if (theSameConfig_Number.Count > 1)
                    {
                        //i 有对应的同一配置，将其添加到同一配置数组的列表中
                        theSameConfigArray_List.Add(theSameConfig_Number.ToArray());
                        //继续下一个 i+1的查询
                        continue;
                    }

                    //剩下的是既不是同一机种，也不是同一配置。
                    theOrther_Number.Add(i);
                }
                #endregion

                //将BaseComprehensive信息写入到workbook的处理
                //创建存放BaseComprehensive个信息的ExcelHelper
                NPOIUse.ExcelHelper excelHelper1 = new NPOIUse.ExcelHelper(excelFileName);
                //同一机种的处理
                if (theSameModelArray_List.Count>0)
                {
                    for (int i = 0; i < theSameModelArray_List.Count; i++)
                    {
                        //获取当前同一机种 组的sheetName
                        string sheetName = baseComprehensiveList[theSameModelArray_List[i][0]].Jobname;

                        //layout  插入的位置,第一个layout插入默认从2开始
                        int layoutInsertRowNum = 2;
                        for (int j = 0; j < theSameModelArray_List[i].Length; j++)
                        {
                            //计算layoutInsertRowNum插入的位置
                            if (j>0)
                            {
                                layoutInsertRowNum += (baseComprehensiveList[theSameModelArray_List[i][j - 1]].feederNozzleDT.Rows.Count + 5);
                            }
                            
                            //将    插入workbook
                            InsertWorkbookfromBaseComprehensive(excelHelper1, sheetName, baseComprehensiveList[theSameModelArray_List[i][j]],
                                2 + j * 6, layoutInsertRowNum);
                            
                        }
                    }
                }

                //同一配置的处理
                if (theSameConfigArray_List.Count>0)
                {
                    for (int i = 0; i < theSameConfigArray_List.Count; i++)
                    {
                        //获取当前同一配置 组的sheetName
                        string sheetName = baseComprehensiveList[theSameConfigArray_List[i][0]].MachineKind;

                        //summaryInfo DataTable的插入位置，第一个默认从2开始
                        int summaryInsertRowNum = 2;
                        //layout  插入的位置,第一个layout插入默认从2开始
                        int layoutInsertRowNum = 2;
                        for (int j = 0; j < theSameConfigArray_List[i].Length; j++)
                        {
                            //计算layoutInsertRowNum插入的位置
                            if (j > 0)
                            {
                                //同一机种job插入位置的指定
                                if (j==1)//处理同一配置 第二个job插入的位置
                                {
                                    summaryInsertRowNum += j * 4;
                                }
                                else//处理同一配置 第三及以上的 job插入的位置
                                {
                                    summaryInsertRowNum += 1;
                                }

                                //layout插入的位置
                                layoutInsertRowNum += (baseComprehensiveList[theSameConfigArray_List[i][j - 1]].feederNozzleDT.Rows.Count + 5);
                            }

                            //将    插入workbook
                            InsertWorkbookfromBaseComprehensive(excelHelper1, sheetName, baseComprehensiveList[theSameConfigArray_List[i][j]],
                                summaryInsertRowNum, layoutInsertRowNum,true);

                        }
                    }
                }

                //其他job的正常处理
                if (theOrther_Number.Count>0)
                {
                    ////layoutFT的插入位置，同样固定即可。为第二行2
                    //layout  插入的位置,第一个layout插入默认从2开始
                    //int layoutInsertRowNum = 2;
                    for (int j = 0; j < theOrther_Number.Count; j++)
                    {
                        //获取当前 组的 不同 baseComprehensive的sheetName
                        string sheetName = "sheet"+(j+1);
                        
                        ////计算layoutInsertRowNum插入的位置
                        //if (j > 0)
                        //{
                        //    layoutInsertRowNum += (baseComprehensiveList[theOrther_Number[j - 1]].feederNozzleDT.Rows.Count + 5);
                        //}

                        //其他job的正常处理 将 BaseComprehensive对象 插入workbook，每次都是单独的sheet，插入位置固定第二行
                        InsertWorkbookfromBaseComprehensive(excelHelper1, sheetName, baseComprehensiveList[theOrther_Number[j]],
                            2 , 2);
                    }
                }

                //ExcelHelper类实现保存的方法
                excelHelper1.SaveWorkbook();

            }


        }

        #region //在指定位置（行）插入baseComprehensive，解决多个baseComprehensive时插入Excel
        //使用参数ExcelHelper类,插入成功返回true
        //summaryInfo插入的位置，LayoutDT插入的位置
        static void InsertWorkbookfromBaseComprehensive(ExcelHelper excelHelper1, string sheetName, BaseComprehensive baseComprehensive,
            int summaryInsertRowNum, int layoutInsertRowNum, bool isTheSameModule = false)
        {
            #region 第一个sheet的获取和生成
            //第一个sheet的标题，以字符串形式插入
            string oneTitle = "FUJI Line Throughput Report " + baseComprehensive.MachineKind;
                        
            //sheetname太长时无法创建对应的sheet。截取前16长度
            if (sheetName.Length>16)
            {
                sheetName = sheetName.Substring(0, 16);
            }
            string sheet1 = sheetName;
                       
            //判断同一机种是否是第一次插入(包含图片、title、DataTable)，是否是第二次插入
            if (isTheSameModule)
            {
                if (excelHelper1.WorkBook.GetSheet(sheet1) == null)
                {
                    //插入sheet对应表格的datatable
                    DataTableToResultSheet1(excelHelper1, baseComprehensive.SummaryInfoDT, sheet1, true,
                        new int[2] { summaryInsertRowNum + 2, 1 });
                    //插入标题
                    StringInsertWorkbook(excelHelper1, oneTitle, sheet1, new int[2] { summaryInsertRowNum, 1 },
                        Orientation.Horizontal, baseComprehensive.SummaryInfoDT.Columns.Count, new int[] { 0, 45 },0,false);

                    //插入图片位置信息
                    //StringInsertWorkbook(excelHelper1, "", sheet1, new int[2] { 2, 1 }, IndexedColors.LightGreen.Index);
                    //Line_short 需要以画布的形式写入，而不是单元格内的文字;
                    //插入图片
                    insertPictureToWorbook(excelHelper1, sheet1, new int[2] { summaryInsertRowNum+1,
                        1 }, baseComprehensive.PictureDataByte, baseComprehensive.ModuleCount,
                        Orientation.Horizontal, baseComprehensive.SummaryInfoDT.Columns.Count, new int[] { 0, 80 });
                }
                else
                {
                    //插入sheet对应表格的datatable
                    //插入行+1，直插入表数据，不插入column name即可
                    DataTableToResultSheet1(excelHelper1, baseComprehensive.SummaryInfoDT, sheet1, false,
                        new int[2] { summaryInsertRowNum, 1 });
                }

            }
            else
            {
                //插入sheet对应表格的datatable
                DataTableToResultSheet1(excelHelper1, baseComprehensive.SummaryInfoDT, sheet1, true, new int[2] { summaryInsertRowNum + 2, 1 });
                //插入标题
                StringInsertWorkbook(excelHelper1, oneTitle, sheet1, new int[2] { summaryInsertRowNum, 1 },
                    Orientation.Horizontal, baseComprehensive.SummaryInfoDT.Columns.Count, new int[] { 0, 45 }, 0,false);

                //插入图片位置信息
                //StringInsertWorkbook(excelHelper1, "", sheet1, new int[2] { 2, 1 }, IndexedColors.LightGreen.Index);
                //Line_short 需要以画布的形式写入，而不是单元格内的文字;
                //插入图片
                insertPictureToWorbook(excelHelper1, sheet1, new int[2] { summaryInsertRowNum+1,
                1 }, baseComprehensive.PictureDataByte, baseComprehensive.ModuleCount,
                    Orientation.Horizontal, baseComprehensive.SummaryInfoDT.Columns.Count, new int[] { 0, 80 });
            }
            #endregion

            #region 生成第二个sheet
            //第二个sheet的标题，以字符串形式插入
            string secondTitle = baseComprehensive.Jobname;
            //第二个sheet的建议
            string ProposalString = "Proposal \n Layout";

            //获取生成的第二个sheet对象在Excel对应的sheetname
            string sheet2 = sheetName + "layout";

            //插入sheet对应表格的datatable
            //插入位置
            int[] layoutInsertPoint = new int[2] { layoutInsertRowNum, 2 };

            //插入layoutDT
            DataTableToResultSheet2(excelHelper1, baseComprehensive.layoutDT, sheet2, true,
                new int[2] { layoutInsertPoint[0] + 1, layoutInsertPoint[1] });
            //插入feederNozzleDT
            DataTableToResultSheet2(excelHelper1, baseComprehensive.feederNozzleDT, sheet2, true,
                new int[2] { layoutInsertPoint[0] + 1, layoutInsertPoint[1] + 10 });
            //插入标题
            StringInsertWorkbook(excelHelper1, secondTitle, sheet2, layoutInsertPoint,
                 Orientation.Horizontal, baseComprehensive.layoutDT.Columns.Count + baseComprehensive.feederNozzleDT.Columns.Count, new int[2] { 0, 50 }, 0);

            //插入建议
            StringInsertWorkbook(excelHelper1, ProposalString, sheet2, new int[2] { layoutInsertPoint[0], layoutInsertPoint[1] - 1 },
                Orientation.Vertical, baseComprehensive.layoutDT.Rows.Count + 2, null, 0);
            //StringInsertWorkbook(excelHelper1, ProposalString, sheet2, new int[2] { 2, 1 }, IndexedColors.BrightGreen.Index);
            #endregion

            //return true;

            //获取对象excelHelper1的文件流，即使用filename创建的文件流，判断是否有文件，没有文件的fa为null
            //FileStream fs = excelHelper1.FS;
            //if (fs == null)
            //{
            //    using (fs = new FileStream(excelFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            //    {
            //        excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
            //        fs.Close();
            //    }

            //}
            //else
            //{
            //    excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
            //    fs.Close();
            //}

            //写入关闭流应该在所有sheet中的DataTable都写入完成以后，
            //InsertWorkbookfromXXX，只是将信息插入excelHelp的WorkbooK
            //excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
            //MessageBox.Show("生成Excel成功！");
            //生成成功，重置两个实例
            //excelHelper = null;
            //excelHelper_another = null;
        }
        #endregion

        #region //插入单个baseComprehensive 到Workbook
        //在指定位置（行）插入baseComprehensive，解决多个baseComprehensive时插入Excel
        //使用参数ExcelHelper类
        static public void InsertWorkbookfromBaseComprehensive(ExcelHelper excelHelper1, string sheetName, BaseComprehensive baseComprehensive)
        {
            //不指定插入行，默认第二行插入。
            InsertWorkbookfromBaseComprehensive(excelHelper1, sheetName, baseComprehensive, 2, 2);
        }
        #endregion

        #region //从生成excelFileName生成Excel，插入BaseComprehensive对应的各个信息到Excel
        static public void GenExclefromBaseComprehensive(string excelFileName, string sheetName,BaseComprehensive baseComprehensive)
        {
            //使用指定的fileName 创建Iworkbook
            if (excelFileName == null)
            {
                MessageBox.Show("未选择任何文件！");
                return;
            }            

            //创建ExcelHelper，用以存放转换后的sheet1和sheet2
            NPOIUse.ExcelHelper excelHelper1 = new NPOIUse.ExcelHelper(excelFileName);

            //调用重载
            InsertWorkbookfromBaseComprehensive(excelHelper1,sheetName, baseComprehensive);

            //将ExcelHelper 的workbook写入到ExcelHelper 对应的文件流
            //ExcelHelper类实现保存的方法
            excelHelper1.SaveWorkbook();
        }
        #endregion

        #region 获取时间戳 
        /// <summary> 
        /// 获取时间戳 
        /// </summary> 
        /// <returns>时间戳字符串</returns> 
        public static Int64 GetTimeStamp()
        {
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);

            return Convert.ToInt64(ts.TotalMilliseconds);
        }
        #endregion

        #region //获取模组的统计字典
        /// <summary>
        /// 获取模组的统计字典
        /// </summary>
        /// <param name="allModuleStrings">所有模组的字符串</param>
        /// <returns></returns>
        public static Dictionary<string, int> GenModuleStasticsDictFromModuleStrings(string[] allModuleStrings)
        {
            //模组统计
            //字典，存放<模组类型:模组个数> //创建统计模组信息的字典
            Dictionary<string, int> module_StatisticsDict = new Dictionary<string, int>();
            for (int j = 0; j < allModuleStrings.Length; j++)
            {
                if (module_StatisticsDict.ContainsKey(allModuleStrings[j]))
                {
                    module_StatisticsDict[allModuleStrings[j]]++;
                }
                else
                {
                    module_StatisticsDict.Add(allModuleStrings[j], 1);
                }
            }
            return module_StatisticsDict;
        }
        #endregion

        #region // 获取Head的统计字典
        ///// <summary>
        ///// 获取Head的统计字典
        ///// </summary>
        ///// <param name="allHeadStrings">所有Head的字符串</param>
        ///// <returns></returns>
        //public static Dictionary<string, int> GenHeadStasticsDictFromHeadStrings(string[] allHeadStrings)
        //{
        //    Dictionary<string, int> Head_Statistics = new Dictionary<string, int>();
        //    for (int j = 0; j < allHeadStrings.Length; j++)
        //    {
        //        // 只有两种情况，包含key和不包含key
        //        if (Head_Statistics.ContainsKey(allHeadStrings[j]))
        //        {
        //            Head_Statistics[allHeadStrings[j]] += 1;
        //        }
        //        else //if (!Head_Statistics.ContainsKey(allHeadTypeString[j]))
        //        {
        //            Head_Statistics.Add(allHeadStrings[j], 1);
        //        }
        //    }
        //    return Head_Statistics;
        //}
        #endregion

        #region //获取拼接的line字符串
        /// <summary>
        /// 获取拼接的line字符串LineString
        /// </summary>
        /// <param name="module_StatisticsDict">模组统计字典</param>
        /// <param name="base_StatisticsDict">base统计字典</param>
        /// <param name="machineKind">机器种类</param>
        /// <returns></returns>
        public static string GenLineStringFromModuleStasticsBaseStastics(Dictionary<string, int> module_StatisticsDict,
            Dictionary<string, int> base_StatisticsDict, string machineKind)
        {
            //"/r/n" 回车换行符
            string Line = "";
            string Line_short = "";
            foreach (var item in module_StatisticsDict)
            {
                Line_short += item.Key + "*" + item.Value.ToString() + "+";
            }
            Line_short = machineKind + "-(" + Line_short.Substring(0, Line_short.Length - 1) + ")";

            //判断4Mhe 2M base的数量，进行拼接
            if (base_StatisticsDict["2MBASE"] == 0)
            {
                Line = Line_short + "\n" + "4MBASE III * " + base_StatisticsDict["4MBASE"];
            }
            if (base_StatisticsDict["2MBASE"] != 0 && base_StatisticsDict["4MBASE"] != 0)
            {
                Line = Line_short + "\n" + "4MBASE III * " + base_StatisticsDict["4MBASE"] + "+2MBASE III * " + base_StatisticsDict["2MBASE"];
            }
            if (base_StatisticsDict["4MBASE"] == 0)
            {
                Line = Line_short + "\n" + "2MBASE III * " + base_StatisticsDict["2MBASE"];
            }
            return Line;
        }
        #endregion

        #region //获取HeadType拼接字符串
        /// <summary>
        /// 获取HeadType拼接字符串
        /// </summary>
        /// <param name="allHeadStrings">所有Head的字符串</param>
        /// <returns></returns>
        public static string GenHeadTypeFormHeadStrings(string[] allHeadStrings)
        {
            Dictionary<string, int> Head_Statistics = new Dictionary<string, int>();
            for (int j = 0; j < allHeadStrings.Length; j++)
            {
                // 只有两种情况，包含key和不包含key
                if (Head_Statistics.ContainsKey(allHeadStrings[j]))
                {
                    Head_Statistics[allHeadStrings[j]] += 1;
                }
                else //if (!Head_Statistics.ContainsKey(allHeadTypeString[j]))
                {
                    Head_Statistics.Add(allHeadStrings[j], 1);
                }
            }

            //获取拼接HeadType
            string Head_Type = string.Empty;
            foreach (var item in Head_Statistics)
            {
                Head_Type += item.Key + "*" + item.Value + "+";
            }
            Head_Type = Head_Type.Substring(0, Head_Type.Length - 1);

            return Head_Type;
        }
        #endregion

        #region // 生成随后的Remark信息
        /// <summary>
        /// 生成随后的Remark信息
        /// </summary>
        /// <param name="module_StatisticsDict">模组统计字典</param>
        /// <param name="base_StatisticsDict">base统计字典</param>
        /// <returns></returns>
        public static string GenRemarkFromModuleBaseStasticsDict(Dictionary<string, int> module_StatisticsDict,
            Dictionary<string, int> base_StatisticsDict)
        {
            return GenRemarkFromModuleBaseStasticsDict( module_StatisticsDict,base_StatisticsDict,null);
        }
        #endregion

        /// <summary>
        /// 重载，加载计算Tray的功率。
        /// </summary>
        /// <param name="module_StatisticsDict"></param>
        /// <param name="base_StatisticsDict"></param>
        /// <param name="allTrayStrings"></param>
        /// <returns></returns>
        public static string GenRemarkFromModuleBaseStasticsDict(Dictionary<string, int> module_StatisticsDict,
            Dictionary<string, int> base_StatisticsDict,string[] allTrayStrings)
        {
            #region //生成Remake信息
            //
            //使用循环获取key的方法，取值。否则有个可能没有相关的key，无法取值报错
            double gonglv = 0.00;
            //Tray盘功率
            if (allTrayStrings!=null&&allTrayStrings.Length>0)
            {
                Dictionary<string, int> tray_StatisticsDict = new Dictionary<string, int>();
                for (int i = 0; i < allTrayStrings.Length; i++)
                {
                    if (tray_StatisticsDict.ContainsKey(allTrayStrings[i]))
                    {
                        tray_StatisticsDict[allTrayStrings[i]]++;
                    }
                    else
                    {
                        tray_StatisticsDict.Add(allTrayStrings[i],1);
                    }
                }
                //加Tray功率
                foreach (var tray_Statistics in tray_StatisticsDict)
                {
                    if (!power_Dictionary.Keys.Contains(tray_Statistics.Key))
                    {
                        continue;
                    }
                    gonglv += tray_StatisticsDict[tray_Statistics.Key] * power_Dictionary[tray_Statistics.Key];
                }
            }

            //加模组功率
            foreach (var module_Statistics in module_StatisticsDict)
            {
                gonglv += module_StatisticsDict[module_Statistics.Key] * power_Dictionary[module_Statistics.Key];
            }
            //加base功率
            gonglv += base_StatisticsDict["2MBASE"] * power_Dictionary["2MBase"] +
                base_StatisticsDict["4MBASE"] * power_Dictionary["4MBase"];

            //计算base耗气量
            int haoqiliang = base_StatisticsDict["4MBASE"] * air_Consumption_Dictionary["4MBase"] +
                base_StatisticsDict["2MBASE"] * air_Consumption_Dictionary["2MBase"];
            //计算需要的ip
            int countu_ip = base_StatisticsDict["2MBASE"] + base_StatisticsDict["4MBASE"];
            //base长度
            double baseLength = ComprehensiveStaticClass.baseLength_Dictionary["2MBase"] * base_StatisticsDict["2MBASE"] +
                ComprehensiveStaticClass.baseLength_Dictionary["4MBase"] * base_StatisticsDict["4MBASE"];
            //最后拼接
            return string.Format("功率：{0}\n耗气量：{1}\n长度：{2}\n当前IP数：{3}",
                gonglv, haoqiliang, baseLength, countu_ip);
            #endregion
        }
    }

    #region //构造一个对象，里面存放
    //jobname(string),TargetConveyor(string),conveyor(string),machineCount,summaryInfoDT摘要的数据表格（DataTable）,
    //图片信息pictureDataByte(byte[]),machineKind summaryInfo表格sheet的标题，
    //AllModuleType 所有module的信息，AllHeadType 所有Head的信息，对比是否为不同配置
    //Jobname  对比不同机种
    public class ExpressSummaryEveryInfo
    {
        private string jobname;
        private string t_or_b;
        private string targetConveyor;
        private string conveyor;
        private int moduleCount;
        private DataTable summaryInfoDT;
        private byte[] pictureDataByte;
        private string machineKind;
        private string[] allModuleType;
        private string[] allHeadType;
        private string[] allTrayStrings;

        public string[] AllTrayStrings
        {
            set
            {
                allTrayStrings = value;
            }
            get
            {
                return allTrayStrings;
            }
        }
        public string T_or_B
        {
            get
            {
                return t_or_b;
            }
            set
            {
                t_or_b = value;
            }
        }
        public string Jobname
        {
            get
            {
                return jobname;
            }
            set
            {
                jobname = value;
            }
        }
        public string TargetConveyor
        {
            set
            {
                targetConveyor = value;
            }
            get
            {
                return targetConveyor;
            }
        }
        public string Conveyor
        {
            set
            {
                conveyor = value;
            }
            get
            {
                return conveyor;
            }
        }
        public int ModuleCount
        {
            set
            {
                moduleCount = value;
            }
            get
            {
                return moduleCount;
            }
        }
        public DataTable SummaryInfoDT
        {
            set
            {
                summaryInfoDT = value;
            }
            get
            {
                return summaryInfoDT;
            }
        }
        public byte[] PictureDataByte
        {
            set
            {
                pictureDataByte = value;
            }
            get
            {
                return pictureDataByte;
            }
        }
        public string MachineKind
        {
            get
            {
                return machineKind;
            }
            set
            {
                machineKind = value;
            }
        }
        public string[] AllModuleType
        {
            set
            {
                allModuleType = value;
            }
            get
            {
                return allModuleType;
            }
        }
        public string[] AllHeadType
        {
            get
            {
                return allHeadType;
            }
            set
            {
                allHeadType = value;
            }
        }
    }

    //新的类，或者construct，继承 SummaryEveryInfo, 在加上layoutExpressInfo中的信息
    /// <summary>
    /// BaseComprehensive类，TimeStamp初始化时赋值，只读属性，不允许在其他时刻修改
    /// </summary>
    public class BaseComprehensive : ExpressSummaryEveryInfo
    {
        private Int64 timeStamp;
        public DataTable layoutDT;
        public DataTable feederNozzleDT;
        public DataTable totalfeederNozzleDT;
        private int baseCount_M_TotalM3;
        public Dictionary<string, int> base_StatisticsDict;       
        public Dictionary<string, int> module_StatisticsDict;       
        
        public List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List = null;

        public string MachineName { get; set; }
        public string CPHRate { get; set; }
        public Int64 TimeStamp
        {
            get
            {
                return timeStamp;
            }
        }

        //timestamp唯一标识 BaseComprehensive 对象，必须要赋值
        public BaseComprehensive(Int64 timeStamp)
        {
            this.timeStamp = timeStamp;
        }

        //综合计算后的 最终的summaryInfoDT
        public void GetTheEndSummaryInfoDT(DataTable expressSummaryInfoDT, string cycletime, string cPH,string cphRate)
        {
            this.GetTheEndSummaryInfoDT(expressSummaryInfoDT, Convert.ToDouble(cycletime), Convert.ToInt32(cPH), cphRate);
        }

        //重载 GetTheEndSummaryInfoDT
        public void GetTheEndSummaryInfoDT(DataTable expressSummaryInfoDT, double cycletime, int cPH,string cphRate)
        {//无 、CPH rate
            //cycletime，board output/hour、board output/day(22H)、 cPH
            expressSummaryInfoDT.Rows[0]["Cycle time\n[((L1+L2)/2) Sec]"] = cycletime;
            expressSummaryInfoDT.Rows[0]["CPH"] = cPH;
            double doubleOutput = Convert.ToDouble(expressSummaryInfoDT.Rows[0]["Board"]) * 3600 / cycletime;
            expressSummaryInfoDT.Rows[0]["Output \nBoard/hour"] = string.Format("{0:0}", doubleOutput);
            expressSummaryInfoDT.Rows[0]["Output\nBoard/Day(22H)"] = string.Format("{0:0}", doubleOutput * 22);
            //由理论cph求得CPH rate
            expressSummaryInfoDT.Rows[0]["CPH Rate"] = cphRate;
            //赋值给InfoDT属性

            // CPP  链接报价单,未实现

            this.SummaryInfoDT = expressSummaryInfoDT;
        }

        /// <summary>
        /// 修改后更新数据的方法
        /// </summary>
        /// <param name="module_Head_Cph_Structs_List"></param>
        /// <param name="base_StatisticsDict"></param>
        /// List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List, 引用类型，已更新过来
        ///Dictionary<string, int> base_StatisticsDict
        public void UpdateForSelfAfterModefy()
        {
            //public Dictionary<string, int> base_StatisticsDict;

            //统计所有Tray的数字，不定长
            List<string> allTrayList = new List<string>();

            //public List<Module_Head_Cph_Struct> module_Head_Cph_Structs_List = null;
            double cphTheory =0.00;
            //更新allmodule head cph layout
            for (int i = 0; i < module_Head_Cph_Structs_List.Count; i++)
            {
                AllModuleType[i] = module_Head_Cph_Structs_List[i].moduleTrayTypestring[0];
                AllHeadType[i] = module_Head_Cph_Structs_List[i].headTypeString_List[0];
                layoutDT.Rows[i][2] = module_Head_Cph_Structs_List[i].moduleTrayTypestring[0];
                layoutDT.Rows[i][3] = module_Head_Cph_Structs_List[i].headTypeString_List[0];

                //提取Tray盘信息
                if (AllModuleType[i].Split('-').Length>1)
                {
                    allTrayList.Add(AllModuleType[i].Split('-')[1]);
                }
                cphTheory += Convert.ToDouble(module_Head_Cph_Structs_List[i].CPH_strings_List[0][0].Split('-')[1]);
            }
            //更新所有的TrayStrings
            AllTrayStrings = allTrayList.ToArray();


            //更新summaryDT中的：
            //  更新Line、HeadType、CPHRate、Remark

            //获取moduleStasticsDict。不更新module_StatisticsDict，因为没有变，只是增加Tray
            //未加入统计和计算Tray的字典，暂时不更新module字典，否则报错
            //Dictionary<string, int> moduleStasticsDict = ComprehensiveStaticClass.GenModuleStasticsDictFromModuleStrings(
            //    this.AllModuleType);

            //更新图片
            PictureDataByte = ComprehensiveStaticClass.GenModulesPicture(AllModuleType, module_StatisticsDict);

            //更新Line
            string Line = ComprehensiveStaticClass.GenLineStringFromModuleStasticsBaseStastics(
                module_StatisticsDict, base_StatisticsDict, MachineKind);
            //获取HeadType
            string headType = ComprehensiveStaticClass.GenHeadTypeFormHeadStrings(AllHeadType);
            //获取CPHRate
            string cphRate = string.Format("{0:P}", Convert.ToDouble(SummaryInfoDT.Rows[0]["CPH"]) / cphTheory);
            //获取Remark
            string remark = ComprehensiveStaticClass.GenRemarkFromModuleBaseStasticsDict(module_StatisticsDict,
                base_StatisticsDict, AllTrayStrings);
            //更新CPHRate
            CPHRate = cphRate;
            //更新summaryDT
            SummaryInfoDT.Rows[0]["Line"] = Line;
            SummaryInfoDT.Rows[0]["Head \nType"] = headType;
            SummaryInfoDT.Rows[0]["CPH Rate"] = cphRate;
            SummaryInfoDT.Rows[0]["Remark"] = remark;
        }
    }

    #region//构造一个对象，里面包含jobName(程式名),单个轨道的layoutExpressDT(layout的表格)，feederNozzleDT信息，TotalfeederNozzleDT信息
    //暂时无法确定是一轨还是二轨，用cycletime可以区别(考虑某个模组不置件情况)
    class layoutExpressEveryInfo
    {
        public string jobName;
        public string t_or_b;
        public DataTable layoutExpressDT;
        public DataTable feederNozzleDT;
        public DataTable totalfeederNozzleDT;
    }
    #endregion
    #endregion

}
