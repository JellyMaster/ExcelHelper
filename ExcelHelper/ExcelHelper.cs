using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class ExcelHelper
    {

        public enum ExcelType
        {
            Xls,
            Xlsx
        }



        public static MemoryStream CreateExcelSheet(DataSet dataToProcess, ExcelType excelType = ExcelType.Xlsx)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                if (dataToProcess != null)
                {

                    switch (excelType)
                    {
                        case ExcelType.Xls:
                            {
                                stream = CreateXlsDocument(dataToProcess);

                                break;

                            }
                        case ExcelType.Xlsx:
                            {
                                stream = CreateXlsxDocument(dataToProcess);
                                break;
                            }
                    }



                }

            }
            catch (Exception error)
            {
                throw error;
            }

            return stream;
        }







        private static MemoryStream CreateXlsxDocument(DataSet dataToProcess)
        {
            MemoryStream stream = new MemoryStream();
            int rowNumber = 1;
            try
            {
                var excelworkbook = new XSSFWorkbook();

                foreach (DataTable table in dataToProcess.Tables)
                {
                    var worksheet = excelworkbook.CreateSheet();

                    var headerRow = worksheet.CreateRow(0);

                    foreach (DataColumn column in table.Columns)
                    {
                        headerRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(column.ColumnName);

                    }

                    //freeze top panel. 
                    worksheet.CreateFreezePane(0, 1, 0, 1);


                    IDataFormat dateformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle dateCellStyle = worksheet.Workbook.CreateCellStyle();
                    dateCellStyle.DataFormat = dateformat.GetFormat("dd MMM yyyy");


                    IDataFormat doubleformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle doubleStyle = worksheet.Workbook.CreateCellStyle();
                    doubleStyle.DataFormat = doubleformat.GetFormat("0.00");

                    IDataFormat intformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle intStyle = worksheet.Workbook.CreateCellStyle();
                    intStyle.DataFormat = intformat.GetFormat("0");

                    foreach (DataRow row in table.Rows)
                    {
                        var sheetRow = worksheet.CreateRow(rowNumber++);

                        foreach (DataColumn column in table.Columns)
                        {
                            bool boolValue = false;
                            string stringValue = row[column].ToString();
                            double doubleValue = 0;
                            IRichTextString richTextValue = null;
                            DateTime dateTimeValue = new DateTime();

                            ICell cell = sheetRow.CreateCell(table.Columns.IndexOf(column));




                            switch (column.DataType.FullName)
                            {
                                case "System.DateTime":
                                    {
                                        if (DateTime.TryParse(stringValue, out dateTimeValue))
                                        {

                                            cell.SetCellValue(dateTimeValue);


                                            cell.CellStyle = dateCellStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }
                                        break;
                                    }
                                case "System.Int32":
                                    {
                                        if (double.TryParse(stringValue, out doubleValue))
                                        {
                                            cell.SetCellValue(doubleValue);

                                            cell.CellStyle = intStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }
                                case "System.Double":
                                    {
                                        if (double.TryParse(stringValue, out doubleValue))
                                        {
                                            cell.SetCellValue(doubleValue);

                                            cell.CellStyle = doubleStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }
                                case "System.Boolean":
                                    {
                                        if (bool.TryParse(stringValue, out boolValue))
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(boolValue);
                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }

                                default:
                                    {

                                        sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        break;
                                    }
                            }







                        }


                    }

                }

                excelworkbook.Write(stream);
            }
            catch (Exception error)
            {
                throw error;
            }

            return stream;
        }



        private static MemoryStream CreateXlsDocument(DataSet dataToProcess)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                var excelworkbook = new HSSFWorkbook();

                foreach (DataTable table in dataToProcess.Tables)
                {
                    var worksheet = excelworkbook.CreateSheet();

                    var headerRow = worksheet.CreateRow(0);

                    foreach (DataColumn column in table.Columns)
                    {
                        headerRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(column.ColumnName);

                    }

                    //freeze top panel. 
                    worksheet.CreateFreezePane(0, 1, 0, 1);


                    IDataFormat dateformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle dateCellStyle = worksheet.Workbook.CreateCellStyle();
                    dateCellStyle.DataFormat = dateformat.GetFormat("dd MMM yyyy");


                    IDataFormat doubleformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle doubleStyle = worksheet.Workbook.CreateCellStyle();
                    doubleStyle.DataFormat = doubleformat.GetFormat("0.00");

                    IDataFormat intformat = worksheet.Workbook.CreateDataFormat();
                    ICellStyle intStyle = worksheet.Workbook.CreateCellStyle();
                    intStyle.DataFormat = intformat.GetFormat("0");



                    int rowNumber = 1;

                    foreach (DataRow row in table.Rows)
                    {
                        var sheetRow = worksheet.CreateRow(rowNumber++);

                        foreach (DataColumn column in table.Columns)
                        {
                            bool boolValue = false;
                            string stringValue = row[column].ToString();
                            double doubleValue = 0;
                            IRichTextString richTextValue = null;
                            DateTime dateTimeValue = new DateTime();


                            ICell cell = sheetRow.CreateCell(table.Columns.IndexOf(column));




                            switch (column.DataType.FullName)
                            {
                                case "System.DateTime":
                                    {
                                        if (DateTime.TryParse(stringValue, out dateTimeValue))
                                        {

                                            cell.SetCellValue(dateTimeValue);


                                            cell.CellStyle = dateCellStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }
                                        break;
                                    }
                                case "System.Int32":
                                    {
                                        if (double.TryParse(stringValue, out doubleValue))
                                        {
                                            cell.SetCellValue(doubleValue);

                                            cell.CellStyle = intStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }
                                case "System.Double":
                                    {
                                        if (double.TryParse(stringValue, out doubleValue))
                                        {
                                            cell.SetCellValue(doubleValue);

                                            cell.CellStyle = doubleStyle;


                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }
                                case "System.Boolean":
                                    {
                                        if (bool.TryParse(stringValue, out boolValue))
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(boolValue);
                                        }
                                        else
                                        {
                                            sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        }

                                        break;
                                    }

                                default:
                                    {

                                        sheetRow.CreateCell(table.Columns.IndexOf(column)).SetCellValue(stringValue);
                                        break;
                                    }
                            }

                        }


                    }

                }

                excelworkbook.Write(stream);
            }
            catch (Exception error)
            {
                throw error;
            }

            return stream;

        }

        public static DataSet CreateDataSetFromExcel(Stream streamToProcess, string fileExtentison = "xlsx")
        {
            DataSet model = new DataSet();




            if (streamToProcess != null)
            {

                if (fileExtentison == "xlsx")
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(streamToProcess);

                    model = ProcessXLSX(workbook);

                }
                else
                {
                    HSSFWorkbook workbook = new HSSFWorkbook(streamToProcess);

                    model = ProcessXLSX(workbook);
                }





            }

            return model;

        }

        private static DataSet ProcessXLSX(HSSFWorkbook workbook)
        {

            DataSet model = new DataSet();
            for (int index = 0; index < workbook.NumberOfSheets; index++)
            {
                ISheet sheet = workbook.GetSheetAt(index);

                if (sheet != null)
                {
                    DataTable table = GenerateTableData(sheet);

                    model.Tables.Add(table);

                }



            }



            return model;

        }

        private static DataTable GenerateTableData(ISheet sheet)
        {

            DataTable table = new DataTable(sheet.SheetName);

            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                //we will assume the first row are the column names 
                IRow row = sheet.GetRow(rowIndex);

                //a completely empty row of data so break out of the process.
                if (row == null)
                {

                    break;
                }

                if (rowIndex == 0)
                {


                    for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                    {
                        string value = row.GetCell(cellIndex).ToString();

                        if (string.IsNullOrEmpty(value))
                        {
                            break;
                        }
                        else
                        {
                            table.Columns.Add(new DataColumn(value));
                        }

                    }
                }
                else
                {
                    //get the data and add to the collection 

                    //now we know the number of columns to iterate through lets get the data and fill up the table. 

                    DataRow datarow = table.NewRow();

                    object[] objectArray = new object[table.Columns.Count];

                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                    {
                        try
                        {
                            ICell cell = row.GetCell(columnIndex);

                            if (cell != null)
                            {
                                objectArray[columnIndex] = cell.ToString();
                            }
                            else
                            {
                                objectArray[columnIndex] = string.Empty;
                            }



                        }
                        catch (Exception error)
                        {
                            Debug.WriteLine(error.Message);
                            Debug.WriteLine("Column Index" + columnIndex);
                            Debug.WriteLine("Row Index" + row.RowNum);
                        }
                    }


                    datarow.ItemArray = objectArray;
                    table.Rows.Add(datarow);

                }



            }

            return table;
        }

        private static DataSet ProcessXLSX(XSSFWorkbook workbook)
        {
            DataSet model = new DataSet();
            for (int index = 0; index < workbook.NumberOfSheets; index++)
            {
                ISheet sheet = workbook.GetSheetAt(index);

                if (sheet != null)
                {
                    DataTable table = GenerateTableData(sheet);

                    model.Tables.Add(table);

                }



            }



            return model;
        }




    }
}
