using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator.Classes
{
    //public static class DataTableExtensionsUnUsed
    //{
    //    ///// <summary>
    //    ///// 'queryOrder' as form: "DEPARTMENT, FIO , DATE_REGISTRATION  ASC"
    //    ///// </summary>
    //    ///// <param name="dataTable"></param>
    //    ///// <param name="queryOrder"></param>
    //    ///// <returns></returns>
    //    //public static DataTable ChangeDataTableScheme(this DataTable dataTable, string queryOrder)
    //    //{
    //    //                DataTable dtExport;

    //    //    // Sort order of collumns
    //    //    using (DataView dv = dataTable.DefaultView)
    //    //    {
    //    //        dv.Sort = queryOrder;
    //    //        dtExport = dv.ToTable();
    //    //    }
    //    //    return dtExport;
    //    //}

    //    //using DocumentFormat.OpenXml.Packaging;
    //    /*
    //            /// <summary>
    //            /// Download DocumentFormat.OpenXml.dll (Microsoft OpenXML)
    //            /// add link to WindowsBase. 
    //            /// add link to lib 'DocumentFormat.OpenXml.dll'. 
    //            /// add using DocumentFormat.OpenXml.Packaging;
    //            /// </summary>
    //            /// <param name="ds"></param>
    //            /// <param name="pathToFile"></param>
    //            public static void ExportToExcelOpenXML(this DataSet ds, string pathToFile, string nameSheet)
    //            {
    //                using (var workbook = SpreadsheetDocument.Create(pathToFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
    //                {
    //                    var workbookPart = workbook.AddWorkbookPart();
    //                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
    //                    {
    //                        Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
    //                    };

    //                    uint sheetId = 1;

    //                    foreach (DataTable table in ds.Tables)
    //                    {
    //                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
    //                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
    //                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

    //                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
    //                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

    //                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
    //                        {
    //                            sheetId =
    //                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
    //                        }

    //                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
    //                        sheets.Append(sheet);

    //                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

    //                        List<String> columns = new List<string>();
    //                        foreach (DataColumn column in table.Columns)
    //                        {
    //                            columns.Add(column.ColumnName);

    //                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
    //                            {
    //                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
    //                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
    //                            };
    //                            headerRow.AppendChild(cell);
    //                        }

    //                        sheetData.AppendChild(headerRow);

    //                        foreach (DataRow dsrow in table.Rows)
    //                        {
    //                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
    //                            foreach (String col in columns)
    //                            {
    //                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
    //                                {
    //                                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
    //                                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
    //                                };
    //                                newRow.AppendChild(cell);
    //                            }

    //                            sheetData.AppendChild(newRow);
    //                        }
    //                    }
    //                }
    //            }

    //            /// <summary>
    //            /// add link to WindowsBase. 
    //            /// add link to lib 'DocumentFormat.OpenXml.dll'. 
    //            /// add using DocumentFormat.OpenXml.Packaging;
    //            /// </summary>
    //            /// <param name="table"></param>
    //            /// <param name="pathToFile"></param>
    //            public static void ExportToExcelOpenXML(this DataTable table, string pathToFile, string nameSheet)
    //            {
    //                using (var workbook = SpreadsheetDocument.Create(pathToFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
    //                {
    //                    var workbookPart = workbook.AddWorkbookPart();
    //                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
    //                    { Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets() };

    //                    uint sheetId = 1;

    //                    {
    //                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
    //                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
    //                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

    //                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
    //                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

    //                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
    //                        {
    //                            sheetId =
    //                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
    //                        }

    //                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
    //                        {
    //                            Id = relationshipId,
    //                            SheetId = sheetId,
    //                            Name = table.TableName
    //                        };

    //                        sheets.Append(sheet);

    //                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

    //                        List<String> columns = new List<string>();
    //                        foreach (DataColumn column in table.Columns)
    //                        {
    //                            columns.Add(column.ColumnName);

    //                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
    //                            {
    //                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
    //                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
    //                            };
    //                            headerRow.AppendChild(cell);
    //                        }

    //                        sheetData.AppendChild(headerRow);

    //                        foreach (DataRow dsrow in table.Rows)
    //                        {
    //                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
    //                            foreach (String col in columns)
    //                            {
    //                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
    //                                {
    //                                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
    //                                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
    //                                };
    //                                newRow.AppendChild(cell);
    //                            }

    //                            sheetData.AppendChild(newRow);
    //                        }
    //                    }
    //                }
    //            }
    //            */

    //    /* add link at Microsoft.Office.Interop.Excel
    //    * and using Excel = Microsoft.Office.Interop.Excel;
    //    * private void ExportDatatableToExcel(DataTable dt, string sufixExportFile) //Заполнение таблицы в Excel  данными
    //     {
    //         _ProgressBar1Start();
    //         int rows = 1;
    //         int rowsInTable = dt.Rows.Count;
    //         int columnsInTable = dt.Columns.Count; // p.Length;

    //         int stepOfProgressCount = (rowsInTable * columnsInTable) / 100;

    //         string lastCell = GetColumnName(columnsInTable) + rowsInTable;
    //         _ProgressWork1Step();
    //         Excel.Application excel = new Excel.Application
    //         {
    //             Visible = false, //делаем объект не видимым
    //             SheetsInNewWorkbook = 1//Количество листов в книге
    //         };

    //         Excel.Workbooks workbooks = excel.Workbooks;
    //         excel.Workbooks.Add(); //добавляем книгу
    //         Excel.Workbook workbook = workbooks[1];
    //         Excel.Sheets sheets = workbook.Worksheets;
    //         Excel.Worksheet sheet = sheets.get_Item(1);
    //         sheet.Name = Path.GetFileNameWithoutExtension(filepathLoadedData);
    //         _ProgressWork1Step();

    //         for (int k = 1; k < columnsInTable; k++)
    //         {
    //             sheet.Cells[k].WrapText = true;
    //             sheet.Cells[1, k].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
    //             sheet.Cells[1, k].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
    //             sheet.Cells[1, k + 1].Value = dt.Columns[k].ColumnName;
    //             //string columnName = dt.Columns[0].Caption;

    //             sheet.Columns[k].Font.Size = 8;
    //             sheet.Columns[k].Font.Name = "Tahoma";

    //             //colourize of collumns
    //             sheet.Cells[1, k].Interior.Color = Color.Silver;
    //             _ProgressWork1Step();
    //         }

    //         //input data and set type of cells - numbers /text
    //         int stepCount = stepOfProgressCount;
    //         foreach (DataRow row in dt.Rows)
    //         {
    //             rows++;
    //             foreach (DataColumn column in dt.Columns)
    //             {
    //                 if (rows > 1)
    //                 {
    //                     if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
    //                     { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
    //                     else
    //                     { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
    //                 }
    //                 sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
    //                 stepCount--;
    //                 if (stepCount == 0)
    //                 {
    //                     _ProgressWork1Step($"Обработано {rows,20 }, строк из {rowsInTable,15}");
    //                     stepCount = stepOfProgressCount;
    //                 }
    //                 //  sheet.Columns[column.Ordinal + 1].AutoFit();
    //             }
    //         }

    //         //Autofilter                
    //         Excel.Range range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];

    //         //ширина колонок - авто
    //         range.Cells.EntireColumn.AutoFit();
    //         _ProgressWork1Step();

    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
    //         range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

    //         range.Select();
    //         _ProgressWork1Step();

    //         range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

    //         workbook.SaveAs(
    //             Path.GetDirectoryName(filepathLoadedData) + @"\" + Path.GetFileNameWithoutExtension(filepathLoadedData) + sufixExportFile,
    //             Excel.XlFileFormat.xlOpenXMLWorkbook,
    //             System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
    //             Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

    //         workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
    //         workbooks.Close();
    //         _ProgressWork1Step(" ");

    //         lastCell = null;
    //         ReleaseObject(range);
    //         ReleaseObject(sheet);
    //         ReleaseObject(sheets);
    //         ReleaseObject(workbook);
    //         ReleaseObject(workbooks);
    //         excel.Quit();
    //         ReleaseObject(excel);
    //         _ProgressBar1Stop();
    //     }*/

    //    /*  private void ExportFullDataTableToExcel() //Заполнение таблицы в Excel всеми данными
    //      {
    //          int rows = 1;
    //          int rowsInTable = dtMobile.Rows.Count;
    //          int columnsInTable = p.Length; // p.Length;
    //          string lastCell = GetColumnName(columnsInTable) + rowsInTable;

    //          Excel.Application excel = new Excel.Application
    //          {
    //              Visible = false, //делаем объект не видимым
    //              SheetsInNewWorkbook = 1//Количество листов в книге
    //          };

    //          Excel.Workbooks workbooks = excel.Workbooks;
    //          excel.Workbooks.Add(); //добавляем книгу
    //          Excel.Workbook workbook = workbooks[1];
    //          Excel.Sheets sheets = workbook.Worksheets;
    //          Excel.Worksheet sheet = sheets.get_Item(1);
    //          sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
    //          // sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    //          HashSet<string> listCollumnsHide = new HashSet<string>(pTranslate);
    //          listCollumnsHide.ExceptWith(new HashSet<string>(pToAccount));

    //          for (int k = 0; k < columnsInTable; k++)
    //          {
    //              sheet.Cells[k + 1].WrapText = true;
    //              sheet.Cells[1, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
    //              sheet.Cells[1, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
    //              sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
    //              sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
    //              sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
    //              sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

    //              sheet.Cells[1, k + 1].Value = pTranslate[k];

    //              sheet.Columns[k + 1].Font.Size = 8;
    //              sheet.Columns[k + 1].Font.Name = "Tahoma";

    //              //colourize of collumns
    //              if (pTranslate[k].Equals("Итого по контракту, грн"))
    //              { sheet.Columns[k + 1].Interior.Color = Color.DarkSeaGreen; }
    //              else if (pTranslate[k].Equals("К оплате владельцем номера, грн"))
    //              { sheet.Columns[k + 1].Interior.Color = Color.SandyBrown; }
    //              else { sheet.Cells[1, k + 1].Interior.Color = Color.Silver; }
    //          }

    //          //input data and set type of cells - numbers /text
    //          foreach (DataRow row in dtMobile.Rows)
    //          {
    //              rows++;
    //              foreach (DataColumn column in dtMobile.Columns)
    //              {
    //                  if (rows == 2)
    //                  {
    //                      if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
    //                      { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
    //                      else
    //                      { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
    //                  }
    //                  sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
    //                  sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
    //                  sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
    //                  sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
    //                  sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
    //                  sheet.Columns[column.Ordinal + 1].AutoFit();
    //              }
    //          }

    //          //Область сортировки   
    //          Excel.Range range = sheet.Range["A2", lastCell];

    //          //По какому столбцу сортировать
    //          string nameColumnSorted = GetColumnName(Array.IndexOf(pTranslate, "Номер телефона абонента") + 1);
    //          Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

    //          //Добавляем параметры сортировки
    //          sheet.Sort.SortFields.Add(rangeKey);
    //          sheet.Sort.SetRange(range);
    //          sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
    //          sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
    //          sheet.Sort.Apply();

    //          //Очищаем фильтр
    //          sheet.Sort.SortFields.Clear();

    //          for (int k = 0; k < pTranslate.Length; k++)
    //          {
    //              foreach (string str in listCollumnsHide)
    //              {
    //                  if (pTranslate[k].Equals(str))
    //                  {
    //                      sheet.Columns[k + 1].Hidden = true;
    //                  }
    //              }
    //          }

    //          //Autofilter                
    //          range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];
    //          range.Select();
    //          range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

    //          workbook.SaveAs(
    //              Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @"_full.xlsx",
    //              Excel.XlFileFormat.xlOpenXMLWorkbook,
    //              System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
    //              Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

    //          workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
    //          workbooks.Close();

    //          listCollumnsHide = null;
    //          nameColumnSorted = null;
    //          lastCell = null;
    //          ReleaseObject(range);
    //          ReleaseObject(rangeKey);
    //          ReleaseObject(sheet);
    //          ReleaseObject(sheets);
    //          ReleaseObject(workbook);
    //          ReleaseObject(workbooks);
    //          excel.Quit();
    //          ReleaseObject(excel);

    //          //  autofill. manualy set number in D1 and D2, then use function
    //          //rng = this.Application.get_Range("D1","D2");
    //          //Excel.Range rng.AutoFill(this.Application.get_Range("D1", "D5"), Excel.XlAutoFillType.xlFillSeries);
    //          //  add comment:
    //          //Excel.Range dateComment = this.Application.get_Range("A1");
    //          //dateComment.AddComment("Comment added " + DateTime.Now.ToString());
    //          //  delete comment:
    //          //if (dateComment.Comment != null) { dateComment.Comment.Delete(); }

    //          // sheet.Cells[1, k + 1].Font.Bold = true;
    //          // (sheet.Cells[1, column.Ordinal + 1] as Microsoft.Office.Interop.Excel.Range).Font.Size = 8;

    //          //объединение ячеек
    //          //sheet.get_Range(sheet.Cells[2, 2], sheet.Cells[4, 4]).Merge(missing);
    //          //(sheet.Columns).ColumnWidth = 15;
    //          // sheet.Columns.Font.Size = Color.LightPink;
    //      }
    //      */
    //    /* private void ExportDataTableToExcelForAccount() //Заполнение таблицы в Excel данными для бухгалтерии
    //     {
    //         int[] pIdxToAccount = new int[]
    //        {
    //             // для бухгалтерии
    //             dtMobile.Columns.IndexOf("Дата счета"),
    //             dtMobile.Columns.IndexOf("Номер телефона абонента"),
    //             dtMobile.Columns.IndexOf("ФИО сотрудника"),
    //             dtMobile.Columns.IndexOf("Затраты по номеру, грн"),
    //             dtMobile.Columns.IndexOf("НДС, грн"),
    //             dtMobile.Columns.IndexOf("ПФ, грн"),
    //             dtMobile.Columns.IndexOf("Итого по контракту, грн"),
    //             dtMobile.Columns.IndexOf("Общая сумма в роуминге, грн"),
    //             dtMobile.Columns.IndexOf("Подразделение"),
    //             dtMobile.Columns.IndexOf("Табельный номер"),
    //             dtMobile.Columns.IndexOf("ТАРИФНАЯ МОДЕЛЬ"),
    //             dtMobile.Columns.IndexOf("К оплате владельцем номера, грн")
    //        };

    //         int rows = 1;
    //         int rowsInTable = dtMobile.Rows.Count;
    //         int columnsInTable = pIdxToAccount.Length; // p.Length;

    //         Excel.Application excel = new Excel.Application
    //         {
    //             Visible = false, //делаем объект не видимым
    //             SheetsInNewWorkbook = 1//Количество листов в книге
    //         };
    //         Excel.Workbooks workbooks = excel.Workbooks;
    //         excel.Workbooks.Add(); //добавляем книгу
    //         Excel.Workbook workbook = workbooks[1];
    //         Excel.Sheets sheets = workbook.Worksheets;
    //         Excel.Worksheet sheet = sheets.get_Item(1);
    //         sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
    //         //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    //         for (int k = 0; k < columnsInTable; k++)
    //         {
    //             sheet.Cells[k + 1].WrapText = true;
    //             sheet.Cells[k + 1].Interior.Color = Color.Silver;
    //             sheet.Cells[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
    //             sheet.Cells[k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
    //             sheet.Cells[1, k + 1].Value = pToAccount[k];
    //             sheet.Columns[k + 1].Font.Size = 8;
    //             sheet.Columns[k + 1].Font.Name = "Tahoma";

    //             switch (k)
    //             {
    //                 case 0:
    //                 case 1:
    //                 case 2:
    //                 case 8:
    //                 case 9:
    //                 case 10:
    //                     {
    //                         sheet.Columns[k + 1].NumberFormat = "@";
    //                         break;
    //                     }
    //                 case 3:
    //                 case 4:
    //                 case 5:
    //                 case 6:
    //                 case 7:
    //                 case 11:
    //                     {
    //                         sheet.Columns[k + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00";
    //                         sheet.Columns[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
    //                         break;
    //                     }
    //             }
    //         }

    //         //colourize of collumns
    //         sheet.Columns[7].Interior.Color = Color.DarkSeaGreen;  //"Итого по контракту, грн"
    //         sheet.Columns[columnsInTable].Interior.Color = Color.SandyBrown;  //"К оплате владельцем номера, грн"

    //         //input data and set type of cells - numbers /text
    //         foreach (DataRow row in dtMobile.Rows)
    //         {
    //             rows++;
    //             for (int column = 0; column < columnsInTable; column++)
    //             {
    //                 sheet.Cells[rows, column + 1].Value = row[pIdxToAccount[column]];
    //             }
    //         }

    //         //Область сортировки          
    //         Excel.Range range = sheet.Range["A2", GetColumnName(columnsInTable) + (rows - 1)];

    //         //По какому столбцу сортировать
    //         string nameColumnSorted = GetColumnName(Array.IndexOf(pIdxToAccount, dtMobile.Columns.IndexOf("Номер телефона абонента")) + 1);
    //         Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

    //         //Добавляем параметры сортировки
    //         sheet.Sort.SortFields.Add(rangeKey);
    //         sheet.Sort.SetRange(range);
    //         sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
    //         sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
    //         sheet.Sort.Apply();
    //         //Очищаем фильтр
    //         sheet.Sort.SortFields.Clear();

    //         //Autofilter
    //         range = sheet.UsedRange; //sheet.Cells.Range["A1", GetColumnName(columnsInTable) + rowsInTable];
    //         range.Select();

    //         //Форматирование колонок (стиль линий обводки)
    //         range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
    //         range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
    //         range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
    //         range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
    //         range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
    //         range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
    //         range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
    //         range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
    //         range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

    //         workbook.SaveAs(Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx",
    //             Excel.XlFileFormat.xlOpenXMLWorkbook,
    //             System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
    //             Excel.XlSaveAsAccessMode.xlExclusive,
    //             System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
    //         workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
    //         workbooks.Close();

    //         ReleaseObject(range);
    //         ReleaseObject(rangeKey);
    //         ReleaseObject(sheet);
    //         ReleaseObject(sheets);
    //         ReleaseObject(workbook);
    //         ReleaseObject(workbooks);
    //         excel.Quit();
    //         ReleaseObject(excel);
    //         MessageBox.Show("Отчет готов и сохранен:" + Environment.NewLine + Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx");
    //     }
         

    //    private void ReleaseObject(object obj)
    //    {
    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
    //    }

    //    static string GetColumnName(int number)
    //    {
    //        string result;
    //        if (number > 0)
    //        {
    //            int alphabets = (number - 1) / 26;
    //            int remainder = (number - 1) % 26;
    //            result = ((char)('A' + remainder)).ToString();
    //            if (alphabets > 0)
    //                result = GetColumnName(alphabets) + result;
    //        }
    //        else
    //            result = null;
    //        return result;
    //    }
    //    */
    //}


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
