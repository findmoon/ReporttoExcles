﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows.Forms;

namespace NPOIUse
{
    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private ISheet sheet=null;
        private FileStream fs = null;
        private bool disposed;

        public IWorkbook WorkBook
        {
            get
            {
                return this.workbook;
            }
        }

        public FileStream FS
        {
            get
            {
                return this.fs;
            }
        }

        public ISheet Sheet
        {
            get
            {
                return sheet;
            }
            set
            {
                sheet = value;
            }
        }

        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            
            createWorkBook();
            disposed = false;
        }

        public void SaveWorkbook()
        {
            if (this.workbook.NumberOfSheets==0)
            {
                MessageBox.Show("没有可生成的数据！");
                return;
            }
            if (this.fileName!=null)
            {
                if (this.fs==null)
                {
                    this.fs = new FileStream(this.fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    //ICSharpCode.SharpZipLib.Zip.ZipException:“EOF in header”，报错应该和没有文件有关
                }

                this.workbook.Write(this.fs);
                this.fs.Close();
                MessageBox.Show("生成Excel成功！");
            }

        }

        //创建workbook
        private void createWorkBook()
        {
            //文件不存在，创建对应的workbook
            if (!File.Exists(this.fileName))
            {
                //this.fs=new FileStream(this.fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //ICSharpCode.SharpZipLib.Zip.ZipException:“EOF in header”，报错应该和没有文件有关
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    this.workbook = new XSSFWorkbook();

                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    this.workbook = new HSSFWorkbook();

            }
            else//存在则读取
            {
                try
                {
                    //FileAccess.ReadWrite、FileAccess.Read若为读写、只读权限，当文档被打开时，下面的操作不会执行，需要给一个“文件已打开”提示
                    //fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    this.fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                    if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    {
                        this.workbook = new XSSFWorkbook(fs);

                    }
                    else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    {
                        this.workbook = new HSSFWorkbook(fs);
                    }
                        
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    this.workbook = null;
                }
            }
            
        }

        /// <summary>
        /// 将DataTable数据导入到Isheet中
        /// </summary>
        /// <param name="data">sheet</param>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="count">写入Isheet的开始的行位置，默认从0行开始</param>
        public static void DataTableToIsheet(ISheet sheet, DataTable dataTable, bool isColumnWritten,int count=0)
        {
            if (dataTable.Rows.Count>=65535)
            {
                return;
            }
            
            int i = 0;
            int j = 0;

            try
            {                
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dataTable.Columns.Count; ++j)
                    {                       
                        row.CreateCell(j).SetCellValue(dataTable.Columns[j].ColumnName);
                    }
                    count += 1;
                }
                

                for (i = 0; i < dataTable.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    foreach (DataColumn dTColumn in dataTable.Columns)
                    {
                        if (dataTable.Rows[i][dTColumn].Equals(DBNull.Value))
                        {//空值
                            continue;
                        }
                        ICell cell = row.CreateCell(dTColumn.Ordinal);
                        //当前表格中的内容
                        string drValue = dataTable.Rows[i][dTColumn].ToString();
                        bool success;
                        switch (dTColumn.DataType.ToString())
                        {
                            case "System.String"://字符串类型
                                cell.SetCellValue(drValue);
                                break;
                            case "System.DateTime"://日期类型
                                DateTime DateV;
                                success = DateTime.TryParse(drValue, out DateV);
                                if (success)
                                {
                                    cell.SetCellValue(DateV);
                                }
                                else
                                {
                                    cell.SetCellValue(drValue);
                                }
                                break;
                            case "System.Boolean"://布尔型
                                bool boolV = false;
                                success = bool.TryParse(drValue, out boolV);
                                if (success)
                                {
                                    cell.SetCellValue(boolV);
                                }
                                else
                                {
                                    cell.SetCellValue(drValue);
                                }
                                break;
                            case "System.Int16"://整型
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                int intV;
                                success = int.TryParse(drValue, out intV);
                                if (success)
                                {
                                    cell.SetCellValue(intV);
                                }
                                else
                                {
                                    cell.SetCellValue(drValue);
                                }
                                break;
                            case "System.Decimal"://浮点型
                            case "System.Double":
                                double doubV;
                                success = double.TryParse(drValue, out doubV);
                                if (success)
                                {
                                    cell.SetCellValue(doubV);
                                }
                                else
                                {
                                    cell.SetCellValue(drValue);
                                }
                                break;
                            case "System.DBNull"://空值处理
                            default:
                                break;
                        }
                        
                    }
                    ++count;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <>
        public void DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            //ISheet sheet = null;
           

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    MessageBox.Show("失败");
                    return;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }

                if (fs==null)
                {
                    fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }
                
                workbook.Write(fs); //写入到excel文件，并关闭流

            }
            catch (Exception ex)
            {
                MessageBox.Show( ex.Message);
                
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            DataTable data = new DataTable();
            int startRow = 0;


            sheet = this.ExcelToIsheet(sheetName);
            ////sheet获取失败
            //if (sheet==null)
            //{
            //    return null;
            //}                    

            try
            {

                if (sheet != null)
                {
                    //sheet没有任何数据时返回null
                    if (sheet.PhysicalNumberOfRows == 0)
                    {
                        MessageBox.Show(string.Format("没有“{0}”sheet或者其为空！", sheetName));
                        return null;
                    }
                    IRow firstRow = sheet.GetRow(0);
                    int firstCellCount = firstRow.LastCellNum;

                    //获取所有的行数
                    //sheet.LastRowNum;获取的是Excel表格中行的索引数，实际行数为sheet.LastRowNum+1
                    //sheet.GetRow(i).LastCellNum,获取时是实际的列数
                    int rowCount = sheet.LastRowNum;

                    //获取列数
                    int cellCount = sheetColumns(sheet);


                    //判断是否有列名，无列名则DataTable无列名
                    if (firstRow == null)
                    {
                        isFirstRowColumn = false;
                    }

                    if (isFirstRowColumn)
                    {
                        //使用行的列数不等的情况
                        //for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        for (int i = 0; i < cellCount; i++)
                        {
                            ICell cell;
                            if (i < firstCellCount)
                            {
                                cell = firstRow.GetCell(i);
                                //DataColumn 的添加需要和GetCell(i)对应，否则如果cell==null，有空行跳过，而Columns.Add还是的顺序添加
                                if (cell != null)
                                {
                                    //string cellValue = cell.StringCellValue;
                                    string cellValue = cell.ToString();
                                    if (cellValue != null)
                                    {
                                        DataColumn column = new DataColumn(cellValue);
                                        data.Columns.Add(column);
                                    }
                                    else
                                    {
                                        MessageBox.Show("获取cell内容失败!");
                                        return null;
                                    }
                                }
                                else
                                {
                                    data.Columns.Add();
                                }
                            }
                            else
                            {
                                //列数不等时，添加列初始化DataTable的结构
                                // add方法column 参数为 null，会引发异常
                                try
                                {
                                    data.Columns.Add();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        //第一行不是列名，但是仍要初始化DataTable的列，否则无法使用NewRow生成到DataRow[j]
                        for (int i = 0; i < cellCount; i++)
                        {
                            try
                            {
                                data.Columns.Add();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                        startRow = sheet.FirstRowNum;
                    }

                    //从startRow 开始转换数据到DataTable
                    for (int i = startRow; i <= rowCount; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null

                        //创建与DataTable一样结构的DataRow
                        DataRow dataRow = data.NewRow();
                        for (int j = 0; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            { //同理，没有数据的单元格都默认是null

                                //.StringCellValue;无法获取到数据，暂不清楚获取的是什么内容
                                //  dataRow[j] = row.GetCell(j).StringCellValue;
                                try
                                {
                                    dataRow[j] = row.GetCell(j).ToString();
                                    //MessageBox.Show(row.GetCell(j).StringCellValue);

                                    /*/---------------------
                                    //创建DataTable，然后使用DataTable的NewRow()方法创建datarow，这样创建出来的DataRow结构和DataTable是一样，但是DataTable创建了没有添加column，所以直接使用datarow[j]无法获取到列，异常退出

                                    dataRow[j] = row.GetCell(j).ToString();
                                    *///----------------------------------------------------------

                                    //dataRow[j] = row.GetCell(j).StringCellValue;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    //throw ex;
                                }

                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public ISheet ExcelToIsheet(){
            return ExcelToIsheet(null);
        }

            /// <summary>
            /// 获取sheetName的sheet
            /// </summary>
            /// <param name="sheetName"></param>
            /// <returns></returns>
        public ISheet ExcelToIsheet(string sheetName)
        {
           // ISheet sheet = null;

            try
            {
                if (sheetName != null)
                {
                    sheet = this.workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        //如果没有sheet，则创建
                        sheet = workbook.CreateSheet(sheetName);

                        //if (workbook.NumberOfSheets == 0)
                        //{
                        //    sheet = workbook.CreateSheet(sheetName);
                        //}
                        //else
                        //{
                        //    sheet = workbook.GetSheetAt(0);
                        //}
                    }
                }
                else//没有指定sheetname时
                {
                    ////如果没有sheet，则创建
                    //if (workbook.NumberOfSheets != 0)
                    //{
                    //    sheet = workbook.GetSheetAt(0);                        
                    //}
                    //else
                    //{
                    //    sheet = workbook.CreateSheet();
                    //}
                    sheet = workbook.CreateSheet();
                }

                return sheet;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        


        /*泛类型行不通
        public ISheet pictureDataToSheet<T>(ISheet sheet, T pictureNPOI,int startRow,int startCol, int endRow,int endCol)
         //   where T: XSSFPicture, HSSFPicture,类型应该只有一种的原因吧，无法执行类型约束为两个类，因为类的约束必须放在第一个
        {

            XSSFPicture pictureNPOI_XSSFPicture = pictureNPOI as XSSFPicture;
            HSSFPalette pictureNPOI_HSSFPalette = pictureNPOI as HSSFPalette;
            //XSSFPicture，HSSFPalette是类，只能有一种类型，正好是泛类型要解决的
            //方法和使用一样，但是T的类型取决类申城的Isheet的类型
            //应该使用重载
            if (true)
            {

            }
            else
            {
                return null;
            }
            workbook.AddPicture(pictureNPOI.)
        }
        */

        /// <summary>
        /// Excel sheet中插入图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="pictureNPOI"></param>
        /// <param name="dx1"></param>
        /// <param name="dy1"></param>
        /// <param name="dx2"></param>
        /// <param name="dy2"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <returns></returns>
        //重载
        public void pictureDataToSheet(ISheet sheet, byte[] pictureNPOI,int dx1,int dy1,int dx2,int dy2, int startRow, int startCol, int endRow, int endCol)
        {
            /*将实际图片转换为pictureData时使用，但是pictureNPOI本身就是picture
            byte[] pictureByte=
            workbook.AddPicture(, PictureType.PNG);
            */
            //判断是否有sheet
            //无，则创建
            if (sheet == null)
            {
                sheet = this.workbook.CreateSheet();
            }

            //执行向sheet写图片
            //创建DrawingPatriarch，存放的容器
            IDrawing patriarch = sheet.CreateDrawingPatriarch();
            ///System.InvalidCastException:“无法将类型为“NPOI.XSSF.UserModel.XSSFDrawing”的对象强制转换为类型“NPOI.HSSF.UserModel.HSSFPatriarch”。”
            ///            HSSFPatriarch patriarch = (HSSFPatriarch)sheetA.CreateDrawingPatriarch();
            ///    根据报错改为如下       
           // IDrawing patriarch = sheet.CreateDrawingPatriarch();

            //XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, startCol, startRow, endCol, endRow);
            IClientAnchor anchor = patriarch.CreateAnchor(dx1, dy1, dx2, dy2, startCol, startRow, endCol, endRow);

            //将图片文件读入workbook，用索引指向该文件
            int pictureIdx = workbook.AddPicture(pictureNPOI, PictureType.PNG);

            //根据读入图片和anchor把图片插到相应的位置
            IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);
            //原始大小显示,重载可指定缩放
            //pict.Resize(0.9);
            //return sheet;
        }
        ////重载
        //public ISheet pictureDataToSheet(ISheet sheet, HSSFPicture pictureNPOI, int startRow, int startCol, int endRow, int endCol)
        //{

        //    workbook.AddPicture(pictureNPOI.)
        //}

        /// <summary>
        /// 画矩形，2.0似乎还未实现，xslx格式未实现
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="content_string"></param>
        /// <param name="dx1"></param>
        /// <param name="dy1"></param>
        /// <param name="dx2"></param>
        /// <param name="dy2"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <returns></returns>
        public ISheet shapJuxingToSheet(ISheet sheet, string content_string, int dx1, int dy1, int dx2, int dy2, int startRow, int startCol, int endRow, int endCol)
        {
            if (sheet == null)
            {
                sheet = this.workbook.CreateSheet();
            }

            
            //创建DrawingPatriarch，存放的容器
            IDrawing patriarch = sheet.CreateDrawingPatriarch();
            //画图形的位置点
            IClientAnchor anchor = patriarch.CreateAnchor(dx1, dy1, dx2, dy2, startCol, startRow, endCol, endRow);
            
            //将图片文件读入workbook，用索引指向该文件
            //int pictureIdx = workbook.AddPicture(pictureNPOI, PictureType.TIFF);
            //IShape recl;

            

            //根据读入图片和anchor把图片插到相应的位置
            //IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);
            //原始大小显示,重载可指定缩放
            //pict.Resize(0.9);
            return sheet;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }

        //静态方法，独立于实例的方法
        public static int sheetColumns(ISheet sheet)
        {
            if (sheet != null)
            {

                //遍历，在列数不等时，确保所有行获取到最大的列数，即列的总数

                //int cellCount = sheet.GetRow(0).LastCellNum;
                //下面循环有判断sheet.GetRow(0)
                int cellCount = 0;
                int rowCount = sheet.LastRowNum;
                for (int i = 0; i <= rowCount; i++)
                {
                    if (sheet.GetRow(i) == null)
                    {
                        continue;
                    }
                    if (cellCount < sheet.GetRow(i).LastCellNum)
                    {
                        cellCount = sheet.GetRow(i).LastCellNum; //一行最后一个cell的编号 即总的列数
                                                                 //MessageBox.Show(i + "行" + sheet.GetRow(i).LastCellNum);
                    }

                }
                return cellCount;
            }
            else
            {
                MessageBox.Show("未指定有效的工作表");
                return 0;
            }
        }

        //重载静态方法
        static public DataTable IsheetToDataTable(ISheet sheet)
        {
            return IsheetToDataTable(sheet, false);
        }

        static public DataTable IsheetToDataTable(ISheet sheet, bool isFirstRowColumn)
        {
            DataTable data = new DataTable();
            int startRow = 0;

            //sheet没有任何数据时返回null
            if (sheet.PhysicalNumberOfRows == 0)
            {
                return null;
            }

            try
            {

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int firstCellCount = firstRow.LastCellNum;

                    //获取所有的行数
                    //sheet.LastRowNum;获取的是Excel表格中行的索引数，实际行数为sheet.LastRowNum+1
                    //sheet.GetRow(i).LastCellNum,获取时是实际的列数
                    int rowCount = sheet.LastRowNum;

                    //获取列数
                    int cellCount = sheetColumns(sheet);


                    //判断是否有列名，无列名则DataTable无列名
                    if (firstRow == null)
                    {
                        isFirstRowColumn = false;
                    }

                    if (isFirstRowColumn)
                    {
                        //使用行的列数不等的情况
                        //for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        for (int i = 0; i < cellCount; ++i)
                        {
                            ICell cell;
                            if (i < firstCellCount)
                            {
                                cell = firstRow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellValue = cell.StringCellValue;
                                    if (cellValue != null)
                                    {
                                        DataColumn column = new DataColumn(cellValue);
                                        data.Columns.Add(column);
                                    }
                                }
                            }
                            else
                            {
                                //列数不等时，添加列初始化DataTable的结构
                                // add方法column 参数为 null，会引发异常

                                try
                                {
                                    data.Columns.Add();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }

                            }



                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        //第一行不是列名，但是仍要初始化DataTable的列，否则无法使用NewRow生成大DataRow[j]
                        for (int i = 0; i < cellCount; i++)
                        {
                            try
                            {
                                data.Columns.Add();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                        startRow = sheet.FirstRowNum;
                    }


                    for (int i = 0; i <= rowCount; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null

                        //创建与DataTable一样结构的DataRow
                        DataRow dataRow = data.NewRow();
                        for (int j = 0; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            { //同理，没有数据的单元格都默认是null

                                //.StringCellValue;无法获取到数据，暂不清楚获取的是什么内容
                                //  dataRow[j] = row.GetCell(j).StringCellValue;
                                try
                                {
                                    dataRow[j] = row.GetCell(j).ToString();
                                    //MessageBox.Show(row.GetCell(j).StringCellValue);

                                    /*/---------------------
                                    //创建DataTable，然后使用DataTable的NewRow()方法创建datarow，这样创建出来的DataRow结构和DataTable是一样，但是DataTable创建了没有添加column，所以直接使用datarow[j]无法获取到列，异常退出

                                    dataRow[j] = row.GetCell(j).ToString();
                                    *///----------------------------------------------------------

                                    //dataRow[j] = row.GetCell(j).StringCellValue;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    //throw ex;
                                }

                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }

    /*是否可以构造一个类，来控制对XSSF和HSSF两个类的约束
    public class  NPOIPocture()
    {

    }
    */
}