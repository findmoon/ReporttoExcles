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
    //枚举类型，代表横向和竖向的枚举类型，用于插入文字时合并单元格方位的选择
    public enum MergeOrientation
    {
        Horizontal,
        Vertical,
        Origin//圆点，表示不合并不横向也不竖向
    }

    //存放一些公共的静态方法、数据（字典）的静态类
    static class ComprehensiveSaticClass
    {
        //电、气每种机器的值是固定的，使用字典，将字典放入结构中
        static public Dictionary<string, double> power_Dictionary = new Dictionary<string, double>()
        {
            {"M3III",0.8 },
            {"M6III",0.9 },
            {"2MBase", 0.5 },
            {"4MBase",0.7 },
            {"LT-Tray",0.5 },
            {"LTC-Tray",0.5 }
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
                "Output\nBoard/Day(22H)", "CPH Rate", "Remark" };

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
        public static byte[] genJobsPicture(string[] allModuleTypeString, Dictionary<string, int> module_Statistics)
        {

            string img_M3 = @".\img\M3.png";
            string img_M6 = @".\img\M6.png";

            if (!File.Exists(img_M3) || !File.Exists(img_M3))
            {
                MessageBox.Show("缺少M3或M6的图片文件，请确认后重试");
                return null;
            }

            //获取两个文件的img
            var imgM3 = Image.FromFile(img_M3);
            var imgM6 = Image.FromFile(img_M6);

            #region//获取要生成的图片的宽高
            //生成最后图片的外边框留白,jiange间隔
            int jiange = 5;
            //int finalWidth = imgM3.Width * module_Statistics[moduleKind3_III]/2 + imgM6.Width*module_Statistics[moduleKind6_III] + jiange*2;

            //统计所需要拼接的图片数量
            int M3Image_count = 0;
            int M6Image_count = 0;
            foreach (var item in module_Statistics)
            {
                //区别keyM3III\M3II\M3I
                if (item.Key.Substring(0, 2) == "M3")
                {
                    M3Image_count += item.Value / 2;
                }
                if (item.Key.Substring(0, 2) == "M6")
                {
                    M6Image_count += item.Value;
                }
            }
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

            for (int i = 0; i < allModuleTypeString.Length; i++)
            {
                //M3模组图像
                if (allModuleTypeString[i].Substring(0, 2) == "M3" && allModuleTypeString[i + 1].Substring(0, 2) == "M3")
                {
                    graph.DrawImage(imgM3, pointX, pointY);
                    pointX += imgM3.Width;
                    i++;
                }
                else if (allModuleTypeString[i].Substring(0, 2) == "M6")//M6模组图像
                {
                    graph.DrawImage(imgM6, pointX, pointY);
                    pointX += imgM6.Width;
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

        

        //将string插入Excel指定位置
        private static void StringInsertExcel(ExcelHelper excelHelper1, string insertString, string sheetName, 
            int[] insertPoint, short colorShort)
        {
            StringInsertExcel(excelHelper1, insertString, sheetName, insertPoint, colorShort, MergeOrientation.Origin, 0);                 
        }

        //重载字符串插入sheet，指定合并插入时的单元格，合并方向和单元格数量
        //MergeOrientation
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="insertString"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint"></param>
        /// <param name="colorShort"></param>
        /// <param name="mergeOrientation"></param>
        /// <param name="mergeNum"></param>
        /// <param name="WidthHeight">插入时的宽高，只实现了设置高度</param>
        private static void StringInsertExcel(ExcelHelper excelHelper1, string insertString, string sheetName,
            int[] insertPoint, short colorShort, MergeOrientation mergeOrientation,int mergeNum,int[] WidthHeight= null)
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

                //合并单元格
                if (mergeOrientation==MergeOrientation.Horizontal)//横向
                {
                    sheet.AddMergedRegion(new CellRangeAddress(insertPoint[0], insertPoint[0], insertPoint[1], insertPoint[1] + mergeNum));
                }
                if (mergeOrientation == MergeOrientation.Vertical)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(insertPoint[0], insertPoint[0] + mergeNum, insertPoint[1], insertPoint[1]));
                }

                #region 样式
                ICellStyle cellStyle = workbook.CreateCellStyle();
                //垂直居中
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //水平对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                //设置开启自动换行,在下面应用到单元格样式时，\n转义字符处理
                cellStyle.WrapText = true;

                //设置背景
                cellStyle.FillBackgroundColor = IndexedColors.White.Index;
                cellStyle.FillForegroundColor = colorShort;
                cellStyle.FillPattern = FillPattern.SolidForeground;

                //设置边框在，在合并的单元格中，边框未应用到整个合并单元格
                //cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
                //cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;

                //应用到单元格样式
                row.GetCell(insertPoint[1]).CellStyle = cellStyle;
                #endregion


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelHelper1">要生成的Excel对应的ExcelHelper类对象</param>
        /// <param name="dt">生成sheet的DataTable</param>
        /// ///////////////////<param name="fileName">生成的文件，在判断fs没有的情况下，创建fs文件流</param>
        /// <param name="sheetNamet">将dt插入到指定sheet</param>
        /// <param name="isColumnWritten">是否吧列名作为单元格插入</param>
        /// <param name="insertPoint">插入的位置</param>
        private static void DataTableToResultSheet1(ExcelHelper excelHelper1, DataTable dt, string sheetName, bool isColumnWritten, int[] insertPoint)
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

                #region sheet1的样式
                //设置列宽
                int commonWidth = 11;
                sheet.SetColumnWidth(0, 1 * 256);
                sheet.SetColumnWidth(1, commonWidth * 256);
                sheet.SetColumnWidth(2, commonWidth * 256);
                sheet.SetColumnWidth(3, System.Convert.ToInt32(System.Convert.ToDouble(commonWidth * 256) * 3));
                sheet.SetColumnWidth(4, commonWidth * 2 * 256);
                sheet.SetColumnWidth(5, commonWidth * 2 * 256);
                sheet.SetColumnWidth(6, commonWidth * 256);
                sheet.SetColumnWidth(7, commonWidth * 256);
                sheet.SetColumnWidth(8, commonWidth * 256);
                sheet.SetColumnWidth(9, commonWidth * 256);
                sheet.SetColumnWidth(10, commonWidth * 256);
                sheet.SetColumnWidth(11, commonWidth * 256);
                sheet.SetColumnWidth(12, commonWidth * 256);

                //合并单元格
                sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, (dt.Columns.Count - 1) + 1));//2.0使用 2.0以下为Region
                sheet.AddMergedRegion(new CellRangeAddress(2, 2, 1, (dt.Columns.Count - 1) + 1));//2.0使用 2.0以下为Region
                
                //设置行高
                IRow row1,
                    row2,
                    row3,
                    row4;
                if (sheet.GetRow(1) == null)
                {
                    row1 = sheet.CreateRow(1);
                }
                else
                {
                    row1 = sheet.GetRow(1);
                }

                if (sheet.GetRow(2) == null)
                {
                    row2 = sheet.CreateRow(2);
                }
                else
                {
                    row2 = sheet.GetRow(2);
                }

                if (sheet.GetRow(3) == null)
                {
                    row3 = sheet.CreateRow(3);
                }
                else
                {
                    row3 = sheet.GetRow(3);
                }

                if (sheet.GetRow(4) == null)
                {
                    row4 = sheet.CreateRow(4);
                }
                else
                {
                    row4 = sheet.GetRow(4);
                }

                row1.HeightInPoints = 50;
                row2.HeightInPoints = 80;
                row3.HeightInPoints = 50;
                row4.HeightInPoints = 50;

                //设置整列单元格样式
                //获取列数量
                //int column_count = ExcelHelper.sheetColumns(sheet);
                //for (int i = 0; i < column_count; i++)
                //{
                //    //设置SetDefaultColumnStyle后，仅到新创建cell时，会应用这个默认设置
                //    //sheet.SetDefaultColumnStyle(i, cellStyle);
                //}
                ICellStyle cellStyle;
                for (int i = 0; i <= sheet.LastRowNum; i++)
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

                            //设置单元格样式,必须每一个cell对饮一个样式，否则无法设置背景

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
                            cellStyle.FillForegroundColor = IndexedColors.LightGreen.Index;

                            //设置边框
                            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

                            
                            if (i == sheet.LastRowNum)
                            {
                                cellStyle.FillForegroundColor = IndexedColors.White.Index;
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

        private static void DataTableToResultSheet2(ExcelHelper excelHelper1, DataTable dt, string sheetNamet, bool isColumnWritten, int[] insertPoint)
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

        private static void setSecondSheetStyle(ISheet sheet, int mergeLastCount)
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

        /// <summary>
        /// 插入图片在ExcelHelper中sheetName表中
        /// </summary>
        /// <param name="excelHelper1"></param>
        /// <param name="sheetName"></param>
        /// <param name="insertPoint">插入位置</param>
        /// <param name="PictureDataByte">图片</param>
        /// <param name="machineCount">根据机器长度判定结束单元格的长短</param>
        private static void insertPictureToExcel(ExcelHelper excelHelper1,string sheetName,int[] insertPoint,
            byte[] PictureDataByte,int machineCount)
        {
            #region 将图片插入sheet的实现
            //生成报告的excelHelper1 excelHelper类
            //IWorkbook workbook = excelHelper1.WorkBook;

            ISheet sheet1 = excelHelper1.ExcelToIsheet(sheetName);

            //从指定的位置开始插入图片
            int count = insertPoint[0];
            //插入的位置
            int startRow = count;
            int startCol = insertPoint[1];
            int endRow = startRow + 1;
            int endCol = startCol + 1;
            //根据模组数量长度，变更结束单元格位置
            
            if (machineCount > 14)
            {
                endCol = startCol + 4;
            }
            if (machineCount > 7)
            {
                endCol = startCol + 3;
            }
            if (machineCount > 3)
            {
                endCol = startCol + 2;
            }

            //偏移依旧不起作用
            excelHelper1.pictureDataToSheet(sheet1, PictureDataByte, 50, 10, 50, 10, startRow, startCol, endRow, endCol);
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
        static public void genExcelfromBaseComprehensiveList(List<BaseComprehensive> baseComprehensiveList,string excelFileName)
        {
            if (baseComprehensiveList.Count==0)
            {
                return;
            }
            if (baseComprehensiveList.Count == 1)//只选择了一个job报告，直接生成
            {
                genExcelfromBaseComprehensive(excelFileName, baseComprehensiveList[0]);                
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

                //同一机种的处理
                if (theSameModelArray_List.Count>0)
                {

                }

                //同一配置的处理
                if (theSameConfigArray_List.Count>0)
                {

                }

                //其他job的正常处理
                if (theOrther_Number.Count>0)
                {

                }
            }


        }

        //在指定位置（行）插入baseComprehensive，解决多个baseComprehensive时插入Excel
        //使用参数ExcelHelper类
        static public void genExcelfromBaseComprehensive(ExcelHelper excelHelper1, BaseComprehensive baseComprehensive,int insertRowNum)
        {

        }

        //插入单个baseComprehensive
        //在指定位置（行）插入baseComprehensive，解决多个baseComprehensive时插入Excel
        //使用参数ExcelHelper类
        static public void genExcelfromBaseComprehensive(ExcelHelper excelHelper1, BaseComprehensive baseComprehensive)
        {

        }

        #region 插入BaseComprehensive对应的各个信息到Excel
        static public void genExcelfromBaseComprehensive(string excelFileName, BaseComprehensive baseComprehensive)
        {
            //使用指定的fileName 创建Iworkbook
            if (excelFileName == null)
            {
                MessageBox.Show("未选择任何文件！");
                return;
            }            

            //创建ExcelHelper，用以存放转换后的sheet1和sheet2
            NPOIUse.ExcelHelper excelHelper1 = new NPOIUse.ExcelHelper(excelFileName);

            #region 第一个sheet的获取和生成
            //第一个sheet的标题，以字符串形式插入
            string oneTitle = "FUJI Line Throughput Report " + baseComprehensive.MachineKind;

            string sheet1 = "sheet1";

            //插入sheet对应表格的datatable
            DataTableToResultSheet1(excelHelper1, baseComprehensive.SummaryInfoDT, sheet1, true, new int[2] { 3, 1 });
            //插入标题
            StringInsertExcel(excelHelper1, oneTitle, sheet1, new int[2] { 1, 1 }, IndexedColors.BrightGreen.Index);

            //插入图片位置信息
            //StringInsertExcel(excelHelper1, "", sheet1, new int[2] { 2, 1 }, IndexedColors.LightGreen.Index);
            //Line_short 需要以画布的形式写入，而不是单元格内的文字;
            //插入图片
            insertPictureToExcel(excelHelper1, sheet1, new int[2] { 2,
                baseComprehensive.SummaryInfoDT.Columns.Count / 2 - 2 },baseComprehensive.PictureDataByte,
                baseComprehensive.MachineCount);
            #endregion

            #region 生成第二个sheet
            //第二个sheet的标题，以字符串形式插入
            string secondTitle = baseComprehensive.Jobname;
            //第二个sheet的建议
            string ProposalString = "Proposal \n Layout";

            //获取生成的第二个sheet对象在Excel对应的sheetname
            string sheet2 = "sheet2";

            //插入sheet对应表格的datatable
            //插入位置
            int[] layoutDTPoint = new int[2] { 3, 2 };
            
            //插入layoutDT
            DataTableToResultSheet2(excelHelper1, baseComprehensive.layoutDT, sheet2, true, layoutDTPoint);
            //插入feederNozzleDT
            DataTableToResultSheet2(excelHelper1, baseComprehensive.feederNozzleDT, sheet2, true, 
                new int[2] { layoutDTPoint[0],layoutDTPoint[1]+10});
            //插入标题
            StringInsertExcel(excelHelper1, secondTitle, sheet2, new int[2] { layoutDTPoint[0]-1, layoutDTPoint[1] }, 
                IndexedColors.BrightGreen.Index, MergeOrientation.Horizontal, 
                baseComprehensive.layoutDT.Columns.Count+ baseComprehensive.feederNozzleDT.Columns.Count-1);

            //插入建议
            StringInsertExcel(excelHelper1, ProposalString, sheet2, new int[2] { 2, 1 }, IndexedColors.BrightGreen.Index,
                MergeOrientation.Vertical, baseComprehensive.layoutDT.Rows.Count+1,new int[2] { 0,50});
            //StringInsertExcel(excelHelper1, ProposalString, sheet2, new int[2] { 2, 1 }, IndexedColors.BrightGreen.Index);
            #endregion


            //获取对象excelHelper1的文件流，即使用filename创建的文件流，判断是否有文件，没有文件的fa为null
            FileStream fs = excelHelper1.FS;
            if (fs == null)
            {
                using (fs = new FileStream(excelFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
                    fs.Close();
                }

            }
            else
            {
                excelHelper1.WorkBook.Write(fs); //写入到excel文件，并关闭流
                fs.Close();
            }


            MessageBox.Show("生成Excel成功！");
            //生成成功，重置两个实例
            //excelHelper = null;
            //excelHelper_another = null;
        }
        #endregion
    }

    //构造一个对象，里面存放
    //jobname(string),TargetConveyor(string),conveyor(string),machineCount,summaryInfoDT摘要的数据表格（DataTable）,
    //图片信息pictureDataByte(byte[]),machineKind summaryInfo表格sheet的标题，
    //AllModuleType 所有module的信息，AllHeadType 所有Head的信息，对比是否为不同配置
    //Jobname  对比不同机种
    class ExpressSummaryEveryInfo
    {
        private string jobname;
        private string t_or_b;
        private string targetConveyor;
        private string conveyor;
        private int machineCount;
        private DataTable summaryInfoDT;
        private byte[] pictureDataByte;
        private string machineKind;
        private string[] allModuleType;
        private string[] allHeadType;

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
        public int MachineCount
        {
            set
            {
                machineCount = value;
            }
            get
            {
                return machineCount;
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
    class BaseComprehensive : ExpressSummaryEveryInfo
    {
        public DataTable layoutDT;
        public DataTable feederNozzleDT;
        public DataTable totalfeederNozzleDT;

        //综合计算后的 最终的summaryInfoDT
        public void GetTheEndSummaryInfoDT(DataTable expressSummaryInfoDT,string cycletime,string cPH)
        {
            this.GetTheEndSummaryInfoDT(expressSummaryInfoDT, Convert.ToDouble(cycletime), Convert.ToInt32(cPH));
        }

        //重载 GetTheEndSummaryInfoDT
        public void GetTheEndSummaryInfoDT(DataTable expressSummaryInfoDT, double cycletime, int cPH)
        {//无 、CPH rate
            //cycletime，board output/hour、board output/day(22H)、 cPH
            expressSummaryInfoDT.Rows[0]["Cycle time\n[((L1+L2)/2) Sec]"] = cycletime;
            expressSummaryInfoDT.Rows[0]["CPH"] = cPH;
            double doubleOutput = Convert.ToDouble(expressSummaryInfoDT.Rows[0]["Board"]) * 3600 / cycletime;
            expressSummaryInfoDT.Rows[0]["Output \nBoard/hour"] = string.Format("{0:0}", doubleOutput);
            expressSummaryInfoDT.Rows[0]["Output\nBoard/Day(22H)"] = string.Format("{0:0}", doubleOutput * 22);
            //由理论cph求得CPH rate

            //赋值给InfoDT属性

            // CPP  链接报价单,未实现

            this.SummaryInfoDT = expressSummaryInfoDT;
        }

        //综合计算后的 最终的LayoutDT
        public void GetTheEndLayoutDT(DataTable expressLayoutDT1,DataTable expressLayoutDT2)
        {
            
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

}
