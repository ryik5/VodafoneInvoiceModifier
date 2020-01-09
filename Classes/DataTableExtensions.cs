using DocumentFormat.OpenXml.Packaging;
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
        public static List<string> ExportRowsToList(this DataTable table)
        {
            List<string> result = new List<string>();

            foreach (DataRow dr in table.Rows)
            {
                result.Add(string.Join("\t", dr.ItemArray));
            }

            return result;
        }

        public static string ExportRowsToText(this DataTable table)
        {
            string result = string.Empty;

            foreach (DataRow dr in table.Rows)
            {
                result += (string.Join("\t", dr.ItemArray) + Environment.NewLine);
            }

            return result;
        }

        public static List<ColumnInfo> ExportColumnInfoToList(this DataTable table)
        {
            List<ColumnInfo> list = new List<ColumnInfo>();

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                list.Add(new ColumnInfo{ColumnName=column.ColumnName, ColumnType=column.DataType.FullName});
            }

            return list;
        }
        public static List<string> ExportColumnNameToList(this DataTable table)
        {
            List<string> list = new List<string>();

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                list .Add(column.Caption);
            }

            return list;
        }

        public static string ExportColumnInfoToText(this DataTable table)
        {
            string result = string.Empty;

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                result += $"Name: {column.ColumnName}\tType: {column.DataType}{Environment.NewLine}";
            }

            return result;
        }
        public static string ExportColumnNameToText(this DataTable table)
        {
            string result = string.Empty;

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                result += $"{column.ColumnName}{Environment.NewLine}";
            }

            return result;
        }

        public static void SetColumnsOrder(this DataTable table, params string[] columnNames)
        {
            List<string> listColNames = columnNames.ToList();
            List<string> listColNamesOfTable = table.ExportColumnNameToList();

            listColNamesOfTable.Except(columnNames.ToList());

            //Remove invalid column names.
            foreach (string colName in columnNames)
            {
                if (!table.Columns.Contains(colName))
                {
                    listColNames.Remove(colName);
                }
            }

            int columnIndex = 0;
            foreach (var columnName in listColNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }

            //foreach (string ColName in listColNamesOfTable)
            //{
            //    if (table.Columns.Contains(ColName))
            //        table.Columns.Remove(ColName);
            //}
        }
        public static void RemoveColumn(this DataTable table, string columnName)
        {
            if (table.Columns.Contains(columnName))
                table.Columns.Remove(columnName);
        }

        /// <summary>
        /// 'queryOrder' as form: "DEPARTMENT, FIO , DATE_REGISTRATION  ASC"
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="queryOrder"></param>
        /// <returns></returns>
        public static DataTable ChangeDataTableScheme(this DataTable dataTable, string queryOrder)
        {
            //  dataTable.SetColumnsOrder(Names.orderColumnsFinacialReport);
            DataTable dtExport;

            // Sort order of collumns
            using (DataView dv = dataTable.DefaultView)
            {
                dv.Sort = queryOrder;
                dtExport = dv.ToTable();
            }
            return dtExport;
        }

        /// <summary>
        /// add link to WindowsBase. 
        /// add link to lib 'DocumentFormat.OpenXml.dll'. 
        /// add using DocumentFormat.OpenXml.Packaging;
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="pathToFile"></param>
        public static void ExportToExcelOpenXML(this DataSet ds, string pathToFile)
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
        public static void ExportToExcelOpenXML(this DataTable table, string pathToFile)
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

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="pathToFile"></param>
        public static void ExportToExcelEPPlus(this DataTable source, string pathToFile)
        {
            DataTable table = source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(pathToFile);
            if (fileInfo.Exists)
                fileInfo.Delete();
            var excel = new ExcelPackage(fileInfo);
            var wsData = excel.Workbook.Worksheets.Add("Data-Worksheetname");
            //var wsPivot = excel.Workbook.Worksheets.Add("Pivot-Worksheetname");
            wsData.Cells["A2"].LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Medium6);
            if (table.Rows.Count != 0)
            {
                foreach (DataColumn col in table.Columns)
                {
                    // format all dates in german format (adjust accordingly)
                    if (col.DataType == typeof(System.DateTime))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 1, colNumber];
                        range.Style.Numberformat.Format = "yyyy.MM.dd"; //"dd.MM.yyyy"
                    }
                }
            }

            var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
            dataRange.AutoFitColumns();
            //var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "Pivotname");
            //pivotTable.MultipleFieldFilters = true;
            //pivotTable.RowGrandTotals = true;
            //pivotTable.ColumGrandTotals = true;
            //pivotTable.Compact = true;
            //pivotTable.CompactData = true;
            //pivotTable.GridDropZones = false;
            //pivotTable.Outline = false;
            //pivotTable.OutlineData = false;
            //pivotTable.ShowError = true;
            //pivotTable.ErrorCaption = "[error]";
            //pivotTable.ShowHeaders = true;
            //pivotTable.UseAutoFormatting = true;
            //pivotTable.ApplyWidthHeightFormats = true;
            //pivotTable.ShowDrill = true;
            //pivotTable.FirstDataCol = 3;
            //pivotTable.RowHeaderCaption = "Claims";

            //var modelField = pivotTable.Fields["Model"];
            //pivotTable.PageFields.Add(modelField);
            //modelField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;

            //var countField = pivotTable.Fields["Claims"];
            //pivotTable.DataFields.Add(countField);

            //var countryField = pivotTable.Fields["Country"];
            //pivotTable.RowFields.Add(countryField);
            //var gspField = pivotTable.Fields["GSP / DRSL"];
            //pivotTable.RowFields.Add(gspField);

            //var oldStatusField = pivotTable.Fields["Old Status"];
            //pivotTable.ColumnFields.Add(oldStatusField);
            //var newStatusField = pivotTable.Fields["New Status"];
            //pivotTable.ColumnFields.Add(newStatusField);

            //var submittedDateField = pivotTable.Fields["Claim Submitted Date"];
            //pivotTable.RowFields.Add(submittedDateField);
            //submittedDateField.AddDateGrouping(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months | OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
            //var monthGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months);
            //monthGroupField.ShowAll = false;
            //var dayGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
            //dayGroupField.ShowAll = false;

            excel.Save();
        }
    }
}
