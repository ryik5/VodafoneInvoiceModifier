//using DocumentFormat.OpenXml.Packaging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public static class DataTableExtensions
    {

        public static List<string> ExportColumnNameToList(this DataTable table)
        {
            List<string> list = new List<string>();

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                list.Add(column.Caption);
            }

            return list;
        }

        public static DataTable SetColumnsOrder(this DataTable table, params string[] columnNames)
        {
            DataTable result = table.Copy();

            List<string> listColNames = columnNames.ToList();
            List<string> listColNamesOfTable = result.ExportColumnNameToList();

            listColNamesOfTable.Except(columnNames.ToList());

            //Remove invalid column names.
            foreach (string colName in columnNames)
            {
                if (!result.Columns.Contains(colName))
                {
                    listColNames.Remove(colName);
                }
            }

            int columnIndex = 0;
            foreach (var columnName in listColNames)
            {
                result.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }

            return result;
        }

        public static DataTable AllowToEditTable(this DataTable table)
        {
            DataTable result = table.Copy();
            foreach (DataColumn col in result.Columns)
            { col.ReadOnly = false; }

            return result;
        }

        /// <summary>
        /// Set DataTable's collection columns on the right order
        /// </summary>
        /// <param name="source">DataTable</param>
        /// <param name="columnsOrder">Columns collection at the right order</param>
        /// <returns>DataTable with changed columns' order and set</returns>
        public static DataTable SeteColumnsCollectionInDataTable(this DataTable source, string[] columnsOrder)
        {
            DataTable result = source;
            List<string> extraColumns = ReturnExtraColumnsInDataTable(result, columnsOrder);
            foreach (var col in extraColumns)
            {
                if (result.Columns.Contains(col))
                    result.Columns.Remove(col);
            }

            return result;
        }

        /// <summary>
        /// Reorder columns' collection in DataTable 
        /// and it will return extra columns' collection which need to delete
        /// </summary>
        /// <param name="source">DataTable which columns' collection is going to change</param>
        /// <param name="columnsAlive">columns' collection which has to put into or stay alived in DataTable</param>
        /// <returns>it will return extra columns' collection which need to delete</returns>
        public static List<string> ReturnExtraColumnsInDataTable(this DataTable source, string[] columnsAlive)
        {
            List<string> columns = source
                .ExportColumnNameToList()
                .Except(columnsAlive.ToList())
                .ToList();
            return columns;
        }

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="pathToFile"></param>
        public static void ExportToExcel(this DataTable source, string pathToFile, string nameSheet, string columnWithColor = null)
        {
            DataTable table = source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(pathToFile);

            if (fileInfo.Exists) fileInfo.Delete();

            //https://riptutorial.com/epplus/example/26056/number-formatting

            var excel = new ExcelPackage(fileInfo);
            var wsData = excel.Workbook.Worksheets.Add(nameSheet);
            wsData.Cells["A2"].LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Medium6);

            if (table.Rows.Count != 0)
            {
                foreach (DataColumn col in table.Columns)
                {
                    // format all dates in german format (adjust accordingly)
                    if (col.DataType == typeof(System.DateTime))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        range.Style.Numberformat.Format = "yyyy.MM.dd"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                    }
                    else if (col.DataType == typeof(System.Decimal) || col.DataType == typeof(System.Double))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        range.Style.Numberformat.Format = "0.00"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                        range.Style.Font.Name = "Tahoma";
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    }
                    else
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        //  range.Style.Numberformat.Format = "@"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                        range.Style.Font.Name = "Tahoma";
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }
                }

                //Set color of special column
                for (int c = 1; c < 1 + table.Columns.Count; c++)
                {
                    if (columnWithColor != null && c == table.Columns.IndexOf(columnWithColor))
                    {
                        for (int r = 3; r < table.Rows.Count + 3; r++)
                        {
                            if (wsData.Cells[r, c + 1]?.ToString()?.Length > 0)
                            {
                                wsData.Cells[r, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                wsData.Cells[r, c + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SandyBrown);
                            }
                        }
                    }
                }

                //Set format of header of table
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.WrapText = true;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.Font.Size = 9;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.Font.Bold = true;
            }

            var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];

            dataRange.AutoFitColumns();

            excel.Save();
        }

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="pathToFile"></param>
        public static void ExportToExcelPivotTable(this DataTable source, string pathToFile, string nameSheet, string[] columnsRedColor = null, string[] columnsGreenColor = null)
        {
            DataTable table = source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(pathToFile);

            if (fileInfo.Exists) fileInfo.Delete();

            //https://riptutorial.com/epplus/example/26056/number-formatting

            var excel = new ExcelPackage(fileInfo);
            var wsData = excel.Workbook.Worksheets.Add(nameSheet);
            var wsPivot = excel.Workbook.Worksheets.Add("Сводная");
            wsData.Cells["A2"].LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Medium6);

            if (table.Rows.Count != 0)
            {
                foreach (DataColumn col in table.Columns)
                {
                    // format all dates in german format (adjust accordingly)
                    if (col.DataType == typeof(System.DateTime))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        range.Style.Numberformat.Format = "yyyy.MM.dd"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                    }
                    else if (col.DataType == typeof(System.Decimal) || col.DataType == typeof(System.Double))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        range.Style.Numberformat.Format = "0.00"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    }
                    else
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                        //  range.Style.Numberformat.Format = "@"; //"dd.MM.yyyy"
                        range.Style.Font.Size = 8;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }
                }

                //Set color of special column
                for (int c = 1; c < 1 + table.Columns.Count; c++)
                {
                    foreach (var col in columnsRedColor)
                    {
                        if (col != null && c == table.Columns.IndexOf(col))
                        {
                            for (int r = 3; r < table.Rows.Count + 3; r++)
                            {
                                if (wsData.Cells[r, c + 1]?.ToString()?.Length > 0)
                                {
                                    wsData.Cells[r, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    wsData.Cells[r, c + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SandyBrown);
                                }
                            }
                        }
                    }
                    foreach (var col in columnsGreenColor)
                    {
                        if (col != null && c == table.Columns.IndexOf(col))
                        {
                            for (int r = 3; r < table.Rows.Count + 3; r++)
                            {
                                if (wsData.Cells[r, c + 1]?.ToString()?.Length > 0)
                                {
                                    wsData.Cells[r, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    wsData.Cells[r, c + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleGreen);
                                }
                            }
                        }
                    }
                }

                //Set format of header of table
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.WrapText = true;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.Font.Size = 9;
                wsData.Cells[2, 1, 2, table.Columns.Count].Style.Font.Bold = true;
            

            var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
            dataRange.Style.Font.Name = "Tahoma";

            dataRange.AutoFitColumns();
            var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "Сводная");
            pivotTable.MultipleFieldFilters = true;
            pivotTable.RowGrandTotals = true;
            pivotTable.ColumGrandTotals = false;
            pivotTable.Compact = true;
            pivotTable.CompactData = true;
            pivotTable.GridDropZones = false;
            pivotTable.Outline = false;
            pivotTable.OutlineData = false;
            pivotTable.ShowError = true;
            pivotTable.ErrorCaption = "[error]";
            pivotTable.ShowHeaders = true;
            pivotTable.UseAutoFormatting = true;
            pivotTable.ApplyWidthHeightFormats = true;
            pivotTable.ShowDrill = true;
            pivotTable.FirstDataCol = 3;
            pivotTable.RowHeaderCaption = "Подразделение";

            var modelField = pivotTable.Fields["ФИО сотрудника"];//Дата счета
            pivotTable.PageFields.Add(modelField);
            modelField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;
            var tarifField = pivotTable.Fields["ТАРИФНАЯ МОДЕЛЬ"];//Дата счета
            pivotTable.PageFields.Add(tarifField);
            tarifField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;
            var numberField = pivotTable.Fields["Номер телефона абонента"];//Дата счета
            pivotTable.PageFields.Add(numberField);
            numberField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;

            var countField = pivotTable.Fields["Итого по контракту, грн"];//Затраты по номеру, грн
            pivotTable.DataFields.Add(countField);
            var paidOwner = pivotTable.Fields["К оплате владельцем номера, грн"];//Затраты по номеру, грн
            pivotTable.DataFields.Add(paidOwner);

            var gspField = pivotTable.Fields["Подразделение"];
            pivotTable.RowFields.Add(gspField);
            gspField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;

            //   var countryField = pivotTable.Fields[""];//Подразделение
            //    pivotTable.RowFields.Add(countryField);

            var oldStatusField = pivotTable.Fields["Дата счета"];//
            pivotTable.ColumnFields.Add(oldStatusField);
                //  var newStatusField = pivotTable.Fields["Общая сумма в роуминге, грн"];
                //  pivotTable.ColumnFields.Add(newStatusField);

                // var submittedDateField = pivotTable.Fields["К оплате владельцем номера, грн"];
                //  pivotTable.RowFields.Add(submittedDateField);
                //   submittedDateField.AddDateGrouping(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months | OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
                //   var monthGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months);
                //   monthGroupField.ShowAll = false;
                //  var dayGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
                //  dayGroupField.ShowAll = false;
            }
            excel.Save();
        }

        //public static List<string> ExportRowsToList(this DataTable table)
        //{
        //    List<string> result = new List<string>();

        //    foreach (DataRow dr in table.Rows)
        //    {
        //        result.Add(string.Join("\t", dr.ItemArray));
        //    }

        //    return result;
        //}

        //public static string ExportRowsToText(this DataTable table)
        //{
        //    string result = string.Empty;

        //    foreach (DataRow dr in table.Rows)
        //    {
        //        result += (string.Join("\t", dr.ItemArray) + Environment.NewLine);
        //    }

        //    return result;
        //}

        //public static List<ColumnInfo> ExportColumnInfoToList(this DataTable table)
        //{
        //    List<ColumnInfo> list = new List<ColumnInfo>();

        //    // Use a DataTable object's DataColumnCollection.
        //    DataColumnCollection columns = table.Columns;

        //    // Print the ColumnName and DataType for each column.
        //    foreach (DataColumn column in columns)
        //    {
        //        list.Add(new ColumnInfo{ColumnName=column.ColumnName, ColumnType=column.DataType.FullName});
        //    }

        //    return list;
        //}


        //public static string ExportColumnInfoToText(this DataTable table)
        //{
        //    string result = string.Empty;

        //    // Use a DataTable object's DataColumnCollection.
        //    DataColumnCollection columns = table.Columns;


        //    foreach (DataColumn column in columns)
        //    {
        //        // Print the ColumnName and DataType for each column.
        //        result += $"Name: {column.ColumnName}\tType: {column.DataType}{Environment.NewLine}";  
        //    }

        //    return result;
        //}

        //public static string ExportColumnNameToText(this DataTable table)
        //{
        //    string result = string.Empty;

        //    // Use a DataTable object's DataColumnCollection.
        //    DataColumnCollection columns = table.Columns;

        //    // Print the ColumnName and DataType for each column.
        //    foreach (DataColumn column in columns)
        //    {
        //        result += $"{column.ColumnName}{Environment.NewLine}";
        //    }

        //    return result;
        //}

        ///// <summary>
        ///// 'queryOrder' as form: "DEPARTMENT, FIO , DATE_REGISTRATION  ASC"
        ///// </summary>
        ///// <param name="dataTable"></param>
        ///// <param name="queryOrder"></param>
        ///// <returns></returns>
        //public static DataTable ChangeDataTableScheme(this DataTable dataTable, string queryOrder)
        //{
        //                DataTable dtExport;

        //    // Sort order of collumns
        //    using (DataView dv = dataTable.DefaultView)
        //    {
        //        dv.Sort = queryOrder;
        //        dtExport = dv.ToTable();
        //    }
        //    return dtExport;
        //}

        /*
                /// <summary>
                /// Download DocumentFormat.OpenXml.dll (Microsoft OpenXML)
                /// add link to WindowsBase. 
                /// add link to lib 'DocumentFormat.OpenXml.dll'. 
                /// add using DocumentFormat.OpenXml.Packaging;
                /// </summary>
                /// <param name="ds"></param>
                /// <param name="pathToFile"></param>
                public static void ExportToExcelOpenXML(this DataSet ds, string pathToFile, string nameSheet)
                {
                    using (var workbook = SpreadsheetDocument.Create(pathToFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        var workbookPart = workbook.AddWorkbookPart();
                        workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
                        {
                            Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
                        };

                        uint sheetId = 1;

                        foreach (DataTable table in ds.Tables)
                        {
                            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                            {
                                sheetId =
                                    sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                            }

                            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                            sheets.Append(sheet);

                            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                            List<String> columns = new List<string>();
                            foreach (DataColumn column in table.Columns)
                            {
                                columns.Add(column.ColumnName);

                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                                {
                                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                                };
                                headerRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(headerRow);

                            foreach (DataRow dsrow in table.Rows)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                                foreach (String col in columns)
                                {
                                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                                    {
                                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
                                    };
                                    newRow.AppendChild(cell);
                                }

                                sheetData.AppendChild(newRow);
                            }
                        }
                    }
                }

                /// <summary>
                /// add link to WindowsBase. 
                /// add link to lib 'DocumentFormat.OpenXml.dll'. 
                /// add using DocumentFormat.OpenXml.Packaging;
                /// </summary>
                /// <param name="table"></param>
                /// <param name="pathToFile"></param>
                public static void ExportToExcelOpenXML(this DataTable table, string pathToFile, string nameSheet)
                {
                    using (var workbook = SpreadsheetDocument.Create(pathToFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        var workbookPart = workbook.AddWorkbookPart();
                        workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
                        { Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets() };

                        uint sheetId = 1;

                        {
                            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                            {
                                sheetId =
                                    sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                            }

                            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                            {
                                Id = relationshipId,
                                SheetId = sheetId,
                                Name = table.TableName
                            };

                            sheets.Append(sheet);

                            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                            List<String> columns = new List<string>();
                            foreach (DataColumn column in table.Columns)
                            {
                                columns.Add(column.ColumnName);

                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                                {
                                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                                };
                                headerRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(headerRow);

                            foreach (DataRow dsrow in table.Rows)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                                foreach (String col in columns)
                                {
                                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                                    {
                                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
                                    };
                                    newRow.AppendChild(cell);
                                }

                                sheetData.AppendChild(newRow);
                            }
                        }
                    }
                }
                */

    }
}
