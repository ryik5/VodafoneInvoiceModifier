using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    public class MakingPivotDataTable : AbstractDataTable, IFilterableDataTable
    {
        ConditionForMakingPivotTable _condition;

        public MakingPivotDataTable() { }

        public MakingPivotDataTable(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            SetDataTable(dataTable);
            SetFilter(condition);
        }

        public void SetFilter(ConditionForMakingPivotTable condition)
        {
            _condition = condition;
        }

        public void SetDataTable(DataTable dataTable)
        {
            Source = dataTable;
        }

        public virtual DataTable MakePivotDataTable2()
        {
            DataTable result = Source
                .AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue).ToString()
                        .Contains(_condition.FilteringService))
                .CopyToDataTable();

            //var results = from myRow in Source.AsEnumerable()
            //              where myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue) == _condition.FilteringService
            //              select myRow;

            return result;
        }

        //DoPivotTable
        public DataTable MakePivotDataTable1()
        {
            DataTable result = Source.AsEnumerable()
                .Where(row => row[_condition.NameColumnWithFilteringServiceValue].ToString().Contains(_condition.FilteringService))
                .GroupBy(row => row[_condition.KeyColumnName]).AsEnumerable()
                .Select(g =>
                {
                    var row = Source.NewRow();
                    DataColumnCollection col = Source.Columns;

                    foreach (DataColumn dc in col)
                    {
                        if (dc.ColumnName.Equals(_condition.KeyColumnName))
                        {
                            row[dc.ColumnName] = g.Key;
                        }
                        else if (dc.ColumnName.Equals(_condition.NameColumnWithFilteringServiceValue))
                        {
                            if ((_condition.TypeResultCalcultedData & TypeData.DataStringMb) == TypeData.DataStringMb)
                            {
                                // Doing as MB ...
                            }
                            else if ((_condition.TypeResultCalcultedData & TypeData.DataStringkB) == TypeData.DataStringkB)
                            {
                                // Doing as kB ...
                            }

                            row[$"{_condition.FilteringService}, Sum"] = g.Sum(r => Int32.Parse(r.Field<string>(dc.ColumnName).ToString()));

                            row[$"{_condition.FilteringService}, Count"] = g.Count();
                        }
                        else
                        {
                            row[dc.ColumnName] = g.Key;
                        }
                    }

                    return row;
                })
                .CopyToDataTable();

            return result;
        }

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="path"></param>
        public void ExportDataTableToPExcelPivot(string path)
        {
            DataTable table = Source;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(path);
            var excel = new ExcelPackage(fileInfo);
            var wsData = excel.Workbook.Worksheets.Add("Data-Worksheetname");
            var wsPivot = excel.Workbook.Worksheets.Add("Pivot-Worksheetname");
            wsData.Cells["A1"].LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Medium6);
            if (table.Rows.Count != 0)
            {
                foreach (DataColumn col in table.Columns)
                {
                    // format all dates in german format (adjust accordingly) 
                    if (col.DataType == typeof(System.DateTime))
                    {
                        var colNumber = col.Ordinal + 1;
                        var range = wsData.Cells[2, colNumber, table.Rows.Count + 1, colNumber];
                        range.Style.Numberformat.Format = "dd.MM.yyyy";
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
            //var gspField = pivotTable.Fields["GSP/DRSL"];
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

    public abstract class AbstractDataTable
    {
        DataTable _source;
        public DataTable Source
        {
            get { return _source.Copy(); }
            set { _source = value; }
        }
    }

    public interface IFilterableDataTable
    {
        DataTable MakePivotDataTable1();
    }

    public class ConditionForMakingPivotTable
    {
        /// <summary>
        /// name of Column which will be used for 'Group by'
        /// </summary>
        public string KeyColumnName { get; set; }

        /// <summary>
        /// service which will be used for 'Filtering'
        /// </summary>
        public string FilteringService { get; set; }

        /// <summary>
        /// name of Column in which the service for 'Filtering' is stored
        /// </summary>
        public string NameColumnWithFilteringService { get; set; }

        /// <summary>
        /// name of Column in which are the calculating data
        /// </summary>
        public string NameColumnWithFilteringServiceValue { get; set; }

        public string FilteringServiceValue { get; set; }
       
        /// <summary>
        /// Type of calculated data
        /// </summary>
        public TypeData TypeResultCalcultedData { get; set; }
    }

    [Flags]
    public enum TypeData
    {
        None = 0,
        DataBool = 1,
        DataInt = 2,
        DataLong = 4,
        DataDouble = 8,
        DataString = 16,
        DataStringMb = 32,
        DataStringkB = 64,
        DataStringB = 128,
    }
}
