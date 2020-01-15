using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System.Collections.Generic;
using System.Data;
using System.Linq;


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
            {
                col.ReadOnly = false;
            }

            return result;
        }

        /// <summary>
        /// Set DataTable's collection columns on the right order
        /// </summary>
        /// <param name="source">DataTable</param>
        /// <param name="columnsOrder">Columns collection at the right order</param>
        /// <returns>DataTable with changed columns' order and set</returns>
        public static DataTable SetColumnsCollectionInDataTable(this DataTable source, string[] columnsOrder)
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
        /// <param name="pathToFile">path to exported file</param>
        /// <param name="nameSheet">name of the sheet</param>
        /// <param name="columnsRedColor">caption columns which data backgroud will be filled red color</param>
        /// <param name="columnsGreenColor">caption columns which data backgroud will be filled green color</param>
        public static void ExportToExcel(this DataTable source, string pathToFile, string nameSheet, string[] columnsRedColor = null, string[] columnsGreenColor = null)
        {
            DataTable table = source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(pathToFile);

            if (fileInfo.Exists) fileInfo.Delete();

            //https://riptutorial.com/epplus/example/26056/number-formatting
            using (ExcelPackage excel = new ExcelPackage(fileInfo))
            {
                var wsData = excel.Workbook.Worksheets.Add(nameSheet);
                wsData.Cells["A2"].LoadFromDataTable(table, true, TableStyles.Medium6);

                if (table.Rows.Count != 0)
                {
                    foreach (DataColumn col in table.Columns)
                    {
                        // format all dates in german format (adjust accordingly)
                        if (col?.DataType == typeof(System.DateTime))
                        {
                            var colNumber = col.Ordinal + 1;
                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                            range.Style.Numberformat.Format = "yyyy.MM.dd"; //"dd.MM.yyyy"
                        }
                        else if (//col?.DataType == typeof(decimal) ||
                            col?.DataType == typeof(int) ||
                            col?.DataType == typeof(long) ||
                            col?.DataType == typeof(double)
                            )
                        {
                            var colNumber = col.Ordinal + 1;
                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                            range.Style.Numberformat.Format = "0.00";
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else
                        {
                            var colNumber = col.Ordinal + 1;
                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                            range.Style.Numberformat.Format = "@"; //"dd.MM.yyyy"
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                    }

                    //Set color cells at special columns
                    for (int c = 1; c < 1 + table.Columns.Count; c++)
                    {
                        if (columnsRedColor?.Length > 0)
                        {
                            foreach (var col in columnsRedColor)
                            {
                                if (col?.Trim().Length > 0 && c == table.Columns.IndexOf(col))
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
                        }
                        if (columnsGreenColor?.Length > 0)
                        {
                            foreach (var col in columnsGreenColor)
                            {
                                if (col?.Trim().Length > 0 && c == table.Columns.IndexOf(col))
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
                    }


                    //Set format of header of table
                    var headerRange = wsData.Cells[2, 1, 2, table.Columns.Count];
                    headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    headerRange.Style.WrapText = true;
                    headerRange.Style.Font.Size = 9;
                    headerRange.Style.Font.Bold = true;

                    //Set format of body of table
                    var bodyRange = wsData.Cells[3, 1, table.Rows.Count + 2, table.Columns.Count];
                    bodyRange.Style.WrapText = false;
                    bodyRange.Style.Font.Size = 8;
                    bodyRange.Style.Font.Bold = false;

                    var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
                    dataRange.Style.Font.Name = "Tahoma";

                    dataRange.AutoFitColumns();
                }
                excel.Save();
            }
        }

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="pathToFile">path to exported file</param>
        /// <param name="nameSheet">name of the sheet</param>
        /// <param name="columnsRedColor">caption columns which data backgroud will be filled red color</param>
        /// <param name="columnsGreenColor">caption columns which data backgroud will be filled green color</param>
        public static void ExportToExcelPivotTable(this DataTable source, string pathToFile, string nameSheet, string[] columnsRedColor = null, string[] columnsGreenColor = null, bool accountant = true)
        {
            DataTable table = source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(pathToFile);

            if (fileInfo.Exists) fileInfo.Delete();

            //https://riptutorial.com/epplus/example/26056/number-formatting
            using (ExcelPackage excel = new ExcelPackage(fileInfo))
            {
                var wsData = excel.Workbook.Worksheets.Add(nameSheet);
                var wsPivot = excel.Workbook.Worksheets.Add("Сводная");
                var wsChart = excel.Workbook.Worksheets.Add("Графики");
                wsData.Cells["A2"].LoadFromDataTable(table, true, TableStyles.Medium6);

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
                        }
                        else if (col.DataType == typeof(decimal) || col.DataType == typeof(int) || col.DataType == typeof(long) || col.DataType == typeof(double))
                        {
                            var colNumber = col.Ordinal + 1;
                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                            range.Style.Numberformat.Format = "0.00"; //"dd.MM.yyyy"
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                        else
                        {
                            var colNumber = col.Ordinal + 1;
                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 2, colNumber];
                            //  range.Style.Numberformat.Format = "@"; //"dd.MM.yyyy"
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                    }

                    //Set color cells at special columns
                    for (int c = 1; c < 1 + table.Columns.Count; c++)
                    {
                        if (columnsRedColor?.Length > 0)
                        {
                            foreach (var col in columnsRedColor)
                            {
                                if (col?.Trim().Length > 0 && c == table.Columns.IndexOf(col))
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
                        }
                        if (columnsGreenColor?.Length > 0)
                        {
                            foreach (var col in columnsGreenColor)
                            {
                                if (col?.Trim().Length > 0 && c == table.Columns.IndexOf(col))
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
                    }

                    //Set format of header of table
                    var headerRange = wsData.Cells[2, 1, 2, table.Columns.Count];
                    headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    headerRange.Style.WrapText = true;
                    headerRange.Style.Font.Size = 9;
                    headerRange.Style.Font.Bold = true;

                    //Set format of body of table
                    var bodyRange = wsData.Cells[3, 1, table.Rows.Count + 2, table.Columns.Count];
                    bodyRange.Style.WrapText = false;
                    bodyRange.Style.Font.Size = 8;
                    bodyRange.Style.Font.Bold = false;

                    //var dataRange1 = wsData.Cells[wsData.Dimension.Address];
                    var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()];
                    dataRange.Style.Font.Name = "Tahoma";
                    dataRange.AutoFitColumns();


                    var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "Сводная");
                    pivotTable.FirstHeaderRow = 1;// first row has headers
                    pivotTable.MultipleFieldFilters = true;
                    pivotTable.RowGrandTotals = true;
                    pivotTable.ColumGrandTotals = true;
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
                    pivotTable.FirstDataCol = 1;// first col of data
                    pivotTable.FirstDataRow = 2;// first row of data

                    if (accountant == true)
                    {
                        pivotTable.Accountant();
                    }
                    else
                    {
                        pivotTable.Market();
                        //      var chart = wsChart.Drawings.AddChart("Графики", OfficeOpenXml.Drawing.Chart.eChartType.BarOfPie, pivotTable);
                        //      chart.SetPosition(1, 0, 1, 0);
                        //       chart.SetSize(300, 300);
                        //      pivotTable.DataOnRows = true; //don't show table
                    }
                }
                excel.Save();
            }
        }

        public static ExcelPivotTable Market(this ExcelPivotTable table)
        {
            var pivotTable = table;

            pivotTable.RowHeaderCaption = "Интернет"; //Имя таблицы

            //Filter
            //  modelField.Sort = eSortType.Ascending;
            //   var pageField =  pivotTable.PageFields.Add(pivotTable.Fields["Имя сервиса"]);
            //  pageField.Sort = eSortType.Ascending;


            //Вычисляемые данные
            var dataField = pivotTable.DataFields.Add(pivotTable.Fields["Суммарно, МБ"]); //"Суммарно, МБ" - имя вычисляемой колонки
            dataField.Name = "Суммарно, МБ";
            dataField.Function = DataFieldFunctions.Sum; //результат - сумма данных в ячейках
            dataField.Field.SubTotalFunctions = eSubTotalFunctions.Sum; //результат - сумма данных в ячейках
            dataField.Format = "0.000";
            dataField.Field.Outline = true;                     //Отображать результат в колонках 
            dataField.Field.ShowInFieldList = true;                 //Отображать результат в колонках

            dataField = pivotTable.DataFields.Add(pivotTable.Fields["Количество"]);
            dataField.Name = "Количество";
            dataField.Function = DataFieldFunctions.Sum; //результат - сумма данных в ячейках
            dataField.Field.SubTotalFunctions = eSubTotalFunctions.Sum;  //результат - сумма данных в ячейках
            dataField.Format = "0";
            dataField.Field.Outline = true;                  //Отображать результат в колонках
            dataField.Field.ShowInFieldList = true;                 //Отображать результат в колонках

            //Rows(Caption)
            var rowField = pivotTable.RowFields.Add(pivotTable.Fields["Подразделение"]);
            rowField.Sort = eSortType.Ascending;
            rowField = pivotTable.RowFields.Add(pivotTable.Fields["ФИО"]);
            rowField.Sort = eSortType.Ascending;
            rowField = pivotTable.RowFields.Add(pivotTable.Fields["Номер телефона"]);
            rowField.Sort = eSortType.Ascending;

            pivotTable.DataOnRows = false;

            return pivotTable;
        }


        public static ExcelPivotTable Accountant(this ExcelPivotTable table)
        {
            var pivotTable = table;

            //Filter
            var modelField = pivotTable.Fields["ФИО сотрудника"];//Дата счета
            pivotTable.PageFields.Add(modelField);
            modelField.Sort = eSortType.Ascending;
            var tarifField = pivotTable.Fields["ТАРИФНАЯ МОДЕЛЬ"];//Дата счета
            pivotTable.PageFields.Add(tarifField);
            tarifField.Sort = eSortType.Ascending;
            var numberField = pivotTable.Fields["Номер телефона абонента"];//Дата счета
            pivotTable.PageFields.Add(numberField);
            numberField.Sort = eSortType.Ascending;

            //Total (Groupby - Calculated values)
            var countField = pivotTable.Fields["Итого по контракту, грн"];//Затраты по номеру, грн
            pivotTable.DataFields.Add(countField);
            var paidOwner = pivotTable.Fields["К оплате владельцем номера, грн"];//Затраты по номеру, грн
            pivotTable.DataFields.Add(paidOwner);

            //Rows(Caption)
            var gspField = pivotTable.Fields["Подразделение"];
            pivotTable.RowFields.Add(gspField);
            gspField.Sort = eSortType.Ascending;

            //   var countryField = pivotTable.Fields[""];//Подразделение
            //    pivotTable.RowFields.Add(countryField);

            //Columns, Total
            var oldStatusField = pivotTable.Fields["Дата счета"];//
            pivotTable.ColumnFields.Add(oldStatusField);
            //  var newStatusField = pivotTable.Fields["Общая сумма в роуминге, грн"];
            //  pivotTable.ColumnFields.Add(newStatusField);

            // var submittedDateField = pivotTable.Fields["К оплате владельцем номера, грн"];
            //  pivotTable.RowFields.Add(submittedDateField);
            //   submittedDateField.AddDateGrouping(eDateGroupBy.Months | eDateGroupBy.Days);
            //   var monthGroupField = pivotTable.Fields.GetDateGroupField(eDateGroupBy.Months);
            //   monthGroupField.ShowAll = false;
            //  var dayGroupField = pivotTable.Fields.GetDateGroupField(eDateGroupBy.Days);
            //  dayGroupField.ShowAll = false;

            return pivotTable;
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

        //using DocumentFormat.OpenXml.Packaging;
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

        /* add link at Microsoft.Office.Interop.Excel
        * and using Excel = Microsoft.Office.Interop.Excel;
        * private void ExportDatatableToExcel(DataTable dt, string sufixExportFile) //Заполнение таблицы в Excel  данными
         {
             _ProgressBar1Start();
             int rows = 1;
             int rowsInTable = dt.Rows.Count;
             int columnsInTable = dt.Columns.Count; // p.Length;

             int stepOfProgressCount = (rowsInTable * columnsInTable) / 100;

             string lastCell = GetColumnName(columnsInTable) + rowsInTable;
             _ProgressWork1Step();
             Excel.Application excel = new Excel.Application
             {
                 Visible = false, //делаем объект не видимым
                 SheetsInNewWorkbook = 1//Количество листов в книге
             };

             Excel.Workbooks workbooks = excel.Workbooks;
             excel.Workbooks.Add(); //добавляем книгу
             Excel.Workbook workbook = workbooks[1];
             Excel.Sheets sheets = workbook.Worksheets;
             Excel.Worksheet sheet = sheets.get_Item(1);
             sheet.Name = Path.GetFileNameWithoutExtension(filepathLoadedData);
             _ProgressWork1Step();

             for (int k = 1; k < columnsInTable; k++)
             {
                 sheet.Cells[k].WrapText = true;
                 sheet.Cells[1, k].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                 sheet.Cells[1, k].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                 sheet.Cells[1, k + 1].Value = dt.Columns[k].ColumnName;
                 //string columnName = dt.Columns[0].Caption;

                 sheet.Columns[k].Font.Size = 8;
                 sheet.Columns[k].Font.Name = "Tahoma";

                 //colourize of collumns
                 sheet.Cells[1, k].Interior.Color = Color.Silver;
                 _ProgressWork1Step();
             }

             //input data and set type of cells - numbers /text
             int stepCount = stepOfProgressCount;
             foreach (DataRow row in dt.Rows)
             {
                 rows++;
                 foreach (DataColumn column in dt.Columns)
                 {
                     if (rows > 1)
                     {
                         if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
                         { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
                         else
                         { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
                     }
                     sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
                     stepCount--;
                     if (stepCount == 0)
                     {
                         _ProgressWork1Step($"Обработано {rows,20 }, строк из {rowsInTable,15}");
                         stepCount = stepOfProgressCount;
                     }
                     //  sheet.Columns[column.Ordinal + 1].AutoFit();
                 }
             }

             //Autofilter                
             Excel.Range range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];

             //ширина колонок - авто
             range.Cells.EntireColumn.AutoFit();
             _ProgressWork1Step();

             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
             range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

             range.Select();
             _ProgressWork1Step();

             range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

             workbook.SaveAs(
                 Path.GetDirectoryName(filepathLoadedData) + @"\" + Path.GetFileNameWithoutExtension(filepathLoadedData) + sufixExportFile,
                 Excel.XlFileFormat.xlOpenXMLWorkbook,
                 System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                 Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

             workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
             workbooks.Close();
             _ProgressWork1Step(" ");

             lastCell = null;
             ReleaseObject(range);
             ReleaseObject(sheet);
             ReleaseObject(sheets);
             ReleaseObject(workbook);
             ReleaseObject(workbooks);
             excel.Quit();
             ReleaseObject(excel);
             _ProgressBar1Stop();
         }*/

        /*  private void ExportFullDataTableToExcel() //Заполнение таблицы в Excel всеми данными
          {
              int rows = 1;
              int rowsInTable = dtMobile.Rows.Count;
              int columnsInTable = p.Length; // p.Length;
              string lastCell = GetColumnName(columnsInTable) + rowsInTable;

              Excel.Application excel = new Excel.Application
              {
                  Visible = false, //делаем объект не видимым
                  SheetsInNewWorkbook = 1//Количество листов в книге
              };

              Excel.Workbooks workbooks = excel.Workbooks;
              excel.Workbooks.Add(); //добавляем книгу
              Excel.Workbook workbook = workbooks[1];
              Excel.Sheets sheets = workbook.Worksheets;
              Excel.Worksheet sheet = sheets.get_Item(1);
              sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
              // sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

              HashSet<string> listCollumnsHide = new HashSet<string>(pTranslate);
              listCollumnsHide.ExceptWith(new HashSet<string>(pToAccount));

              for (int k = 0; k < columnsInTable; k++)
              {
                  sheet.Cells[k + 1].WrapText = true;
                  sheet.Cells[1, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                  sheet.Cells[1, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                  sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                  sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                  sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                  sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                  sheet.Cells[1, k + 1].Value = pTranslate[k];

                  sheet.Columns[k + 1].Font.Size = 8;
                  sheet.Columns[k + 1].Font.Name = "Tahoma";

                  //colourize of collumns
                  if (pTranslate[k].Equals("Итого по контракту, грн"))
                  { sheet.Columns[k + 1].Interior.Color = Color.DarkSeaGreen; }
                  else if (pTranslate[k].Equals("К оплате владельцем номера, грн"))
                  { sheet.Columns[k + 1].Interior.Color = Color.SandyBrown; }
                  else { sheet.Cells[1, k + 1].Interior.Color = Color.Silver; }
              }

              //input data and set type of cells - numbers /text
              foreach (DataRow row in dtMobile.Rows)
              {
                  rows++;
                  foreach (DataColumn column in dtMobile.Columns)
                  {
                      if (rows == 2)
                      {
                          if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
                          { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
                          else
                          { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
                      }
                      sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
                      sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                      sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                      sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                      sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                      sheet.Columns[column.Ordinal + 1].AutoFit();
                  }
              }

              //Область сортировки   
              Excel.Range range = sheet.Range["A2", lastCell];

              //По какому столбцу сортировать
              string nameColumnSorted = GetColumnName(Array.IndexOf(pTranslate, "Номер телефона абонента") + 1);
              Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

              //Добавляем параметры сортировки
              sheet.Sort.SortFields.Add(rangeKey);
              sheet.Sort.SetRange(range);
              sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
              sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
              sheet.Sort.Apply();

              //Очищаем фильтр
              sheet.Sort.SortFields.Clear();

              for (int k = 0; k < pTranslate.Length; k++)
              {
                  foreach (string str in listCollumnsHide)
                  {
                      if (pTranslate[k].Equals(str))
                      {
                          sheet.Columns[k + 1].Hidden = true;
                      }
                  }
              }

              //Autofilter                
              range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];
              range.Select();
              range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

              workbook.SaveAs(
                  Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @"_full.xlsx",
                  Excel.XlFileFormat.xlOpenXMLWorkbook,
                  System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                  Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

              workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
              workbooks.Close();

              listCollumnsHide = null;
              nameColumnSorted = null;
              lastCell = null;
              ReleaseObject(range);
              ReleaseObject(rangeKey);
              ReleaseObject(sheet);
              ReleaseObject(sheets);
              ReleaseObject(workbook);
              ReleaseObject(workbooks);
              excel.Quit();
              ReleaseObject(excel);

              //  autofill. manualy set number in D1 and D2, then use function
              //rng = this.Application.get_Range("D1","D2");
              //Excel.Range rng.AutoFill(this.Application.get_Range("D1", "D5"), Excel.XlAutoFillType.xlFillSeries);
              //  add comment:
              //Excel.Range dateComment = this.Application.get_Range("A1");
              //dateComment.AddComment("Comment added " + DateTime.Now.ToString());
              //  delete comment:
              //if (dateComment.Comment != null) { dateComment.Comment.Delete(); }

              // sheet.Cells[1, k + 1].Font.Bold = true;
              // (sheet.Cells[1, column.Ordinal + 1] as Microsoft.Office.Interop.Excel.Range).Font.Size = 8;

              //объединение ячеек
              //sheet.get_Range(sheet.Cells[2, 2], sheet.Cells[4, 4]).Merge(missing);
              //(sheet.Columns).ColumnWidth = 15;
              // sheet.Columns.Font.Size = Color.LightPink;
          }
          */
        /* private void ExportDataTableToExcelForAccount() //Заполнение таблицы в Excel данными для бухгалтерии
         {
             int[] pIdxToAccount = new int[]
            {
                 // для бухгалтерии
                 dtMobile.Columns.IndexOf("Дата счета"),
                 dtMobile.Columns.IndexOf("Номер телефона абонента"),
                 dtMobile.Columns.IndexOf("ФИО сотрудника"),
                 dtMobile.Columns.IndexOf("Затраты по номеру, грн"),
                 dtMobile.Columns.IndexOf("НДС, грн"),
                 dtMobile.Columns.IndexOf("ПФ, грн"),
                 dtMobile.Columns.IndexOf("Итого по контракту, грн"),
                 dtMobile.Columns.IndexOf("Общая сумма в роуминге, грн"),
                 dtMobile.Columns.IndexOf("Подразделение"),
                 dtMobile.Columns.IndexOf("Табельный номер"),
                 dtMobile.Columns.IndexOf("ТАРИФНАЯ МОДЕЛЬ"),
                 dtMobile.Columns.IndexOf("К оплате владельцем номера, грн")
            };

             int rows = 1;
             int rowsInTable = dtMobile.Rows.Count;
             int columnsInTable = pIdxToAccount.Length; // p.Length;

             Excel.Application excel = new Excel.Application
             {
                 Visible = false, //делаем объект не видимым
                 SheetsInNewWorkbook = 1//Количество листов в книге
             };
             Excel.Workbooks workbooks = excel.Workbooks;
             excel.Workbooks.Add(); //добавляем книгу
             Excel.Workbook workbook = workbooks[1];
             Excel.Sheets sheets = workbook.Worksheets;
             Excel.Worksheet sheet = sheets.get_Item(1);
             sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
             //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

             for (int k = 0; k < columnsInTable; k++)
             {
                 sheet.Cells[k + 1].WrapText = true;
                 sheet.Cells[k + 1].Interior.Color = Color.Silver;
                 sheet.Cells[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                 sheet.Cells[k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                 sheet.Cells[1, k + 1].Value = pToAccount[k];
                 sheet.Columns[k + 1].Font.Size = 8;
                 sheet.Columns[k + 1].Font.Name = "Tahoma";

                 switch (k)
                 {
                     case 0:
                     case 1:
                     case 2:
                     case 8:
                     case 9:
                     case 10:
                         {
                             sheet.Columns[k + 1].NumberFormat = "@";
                             break;
                         }
                     case 3:
                     case 4:
                     case 5:
                     case 6:
                     case 7:
                     case 11:
                         {
                             sheet.Columns[k + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00";
                             sheet.Columns[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                             break;
                         }
                 }
             }

             //colourize of collumns
             sheet.Columns[7].Interior.Color = Color.DarkSeaGreen;  //"Итого по контракту, грн"
             sheet.Columns[columnsInTable].Interior.Color = Color.SandyBrown;  //"К оплате владельцем номера, грн"

             //input data and set type of cells - numbers /text
             foreach (DataRow row in dtMobile.Rows)
             {
                 rows++;
                 for (int column = 0; column < columnsInTable; column++)
                 {
                     sheet.Cells[rows, column + 1].Value = row[pIdxToAccount[column]];
                 }
             }

             //Область сортировки          
             Excel.Range range = sheet.Range["A2", GetColumnName(columnsInTable) + (rows - 1)];

             //По какому столбцу сортировать
             string nameColumnSorted = GetColumnName(Array.IndexOf(pIdxToAccount, dtMobile.Columns.IndexOf("Номер телефона абонента")) + 1);
             Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

             //Добавляем параметры сортировки
             sheet.Sort.SortFields.Add(rangeKey);
             sheet.Sort.SetRange(range);
             sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
             sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
             sheet.Sort.Apply();
             //Очищаем фильтр
             sheet.Sort.SortFields.Clear();

             //Autofilter
             range = sheet.UsedRange; //sheet.Cells.Range["A1", GetColumnName(columnsInTable) + rowsInTable];
             range.Select();

             //Форматирование колонок (стиль линий обводки)
             range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
             range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
             range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
             range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
             range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
             range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
             range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
             range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
             range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

             workbook.SaveAs(Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx",
                 Excel.XlFileFormat.xlOpenXMLWorkbook,
                 System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                 Excel.XlSaveAsAccessMode.xlExclusive,
                 System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
             workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
             workbooks.Close();

             ReleaseObject(range);
             ReleaseObject(rangeKey);
             ReleaseObject(sheet);
             ReleaseObject(sheets);
             ReleaseObject(workbook);
             ReleaseObject(workbooks);
             excel.Quit();
             ReleaseObject(excel);
             MessageBox.Show("Отчет готов и сохранен:" + Environment.NewLine + Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx");
         }
         

        private void ReleaseObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        }

        static string GetColumnName(int number)
        {
            string result;
            if (number > 0)
            {
                int alphabets = (number - 1) / 26;
                int remainder = (number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                    result = GetColumnName(alphabets) + result;
            }
            else
                result = null;
            return result;
        }
        */
    }


    //public interface IPivotTableCreator
    //{
    //    void CreatePivotTable(
    //        OfficeOpenXml.ExcelPackage pkg, // reference to the destination book
    //        string tableName,               // "tab" name used to generate names for related items
    //        string pivotRangeName);         // Named range in the Workbook refers to data
    //}
    //public class SimplePivotTable : IPivotTableCreator
    //{
    //    List<string> _GroupByColumns;
    //    List<string> _SummaryColumns;
    //    /// <summary>
    //    /// Constructor
    //    /// </summary>
    //    public SimplePivotTable(string[] groupByColumns, string[] summaryColumns)
    //    {
    //        _GroupByColumns = new List<string>(groupByColumns);
    //        _SummaryColumns = new List<string>(summaryColumns);
    //    }

    //    /// <summary>
    //    /// Call-back handler that builds simple PivatTable in Excel
    //    /// http://stackoverflow.com/questions/11650080/epplus-pivot-tables-charts
    //    /// </summary>
    //    public void CreatePivotTable(OfficeOpenXml.ExcelPackage pkg, string tableName, string pivotRangeName)
    //    {
    //        string pageName = "Pivot-" + tableName.Replace(" ", "");
    //        var wsPivot = pkg.Workbook.Worksheets.Add(pageName);
    //        pkg.Workbook.Worksheets.MoveBefore(pageName, tableName);

    //        var dataRange = pkg.Workbook./*Worksheets[tableName].*/Names[pivotRangeName];
    //        var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["C3"], dataRange, "Pivot_" + tableName.Replace(" ", ""));
    //        pivotTable.ShowHeaders = true;
    //        pivotTable.UseAutoFormatting = true;
    //        pivotTable.ApplyWidthHeightFormats = true;
    //        pivotTable.ShowDrill = true;
    //        pivotTable.FirstHeaderRow = 1;  // first row has headers
    //        pivotTable.FirstDataCol = 1;    // first col of data
    //        pivotTable.FirstDataRow = 2;    // first row of data

    //        foreach (string row in _GroupByColumns)
    //        {
    //            var field = pivotTable.Fields[row];
    //            pivotTable.RowFields.Add(field);
    //            field.Sort = eSortType.Ascending;
    //        }

    //        foreach (string column in _SummaryColumns)
    //        {
    //            var field = pivotTable.Fields[column];
    //            ExcelPivotTableDataField result = pivotTable.DataFields.Add(field);
    //        }

    //        pivotTable.DataOnRows = false;
    //    }
    //}

}
