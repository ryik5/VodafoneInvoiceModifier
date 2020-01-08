using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    public class MakingPivotDataTable : AbstractDataTable, IFilterableDataTable
    {
        public delegate void MessageStatus(object sender, TextEventArgs e);

        public event MessageStatus Status;

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
            Source = dataTable.Copy();
            Source.Columns.Add("Результат", typeof(decimal));
        }

        private DataColumnCollection SourceDataTableInfo(string calledByMethod)
        {
            Status?.Invoke(this, new TextEventArgs($"Calling method: {calledByMethod}"));
            Status?.Invoke(this, new TextEventArgs($"Колонок в таблице: {Source.Columns.Count}"));

            //List<string> columnsName = new List<string>();
            DataColumnCollection columns = Source.Columns;
            foreach (DataColumn dc in columns)
            {
                //columnsName.Add(dc.ColumnName);
            }
            //Status?.Invoke(this, new TextEventArgs($"Колонки в таблице:{Environment.NewLine}{columnsName.AsString(Environment.NewLine)}"));
            
            Status?.Invoke(this, new TextEventArgs($"Строк в таблице: {Source.Rows.Count}"));

            return columns;
        }

        public virtual DataTable MakePivotDataTable4()
        {
            SourceDataTableInfo(nameof(MakePivotDataTable4));

            var result = Source.AsEnumerable()
                        .GroupBy(r => new { Col1 = r[_condition.KeyColumnName], Col2 = _condition.GroupByOrderColumns })
                        .Select(g => g.OrderBy(x => x.ItemArray = _condition.GroupByOrderColumns).First())
                        .CopyToDataTable();
            return result;
        }


        public virtual DataTable MakePivotDataTable2()
        {
            SourceDataTableInfo(nameof(MakePivotDataTable2));

            return Source.AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringService)
                        .Contains(_condition.FilteringService))
                ?.CopyToDataTable();
        }

        //Do PivotTable
        public virtual DataTable MakePivotDataTable1()
        {
            SourceDataTableInfo(nameof(MakePivotDataTable1));

            var result1 =
                from student in Source.AsEnumerable()
                group student
                by new { Col1 = student.Field<string>(_condition.KeyColumnName) };


            var result = Source.AsEnumerable()
                        .GroupBy(r => new { Col1 = r[_condition.KeyColumnName], Col2 = _condition.GroupByOrderColumns })
                        .Select(g => g.OrderBy(x => x.ItemArray = _condition.GroupByOrderColumns).First())
                        ;

            //var result = Source.AsEnumerable()
            //    .Select(a => new
            //    {
            //        keyColumn = a.Field<String>(_condition.KeyColumnName),
            //        filteringService = a.Field<String>(_condition.NameColumnWithFilteringService),
            //        filteringServiceValue = a.Field<String>(_condition.NameColumnWithFilteringServiceValue),
            //        Value = a.Field<String>(_condition.NameColumnWithFilteringServiceValue).TryParseAsInternetTrafic("Mb"),
            //    })
            //    .GroupBy(r => new { r.keyColumn, r.filteringServiceValue, r.filteringService, r.Value })
            //    .Select(g =>
            //        {
            //            var row = Source.NewRow();

            //            row[_condition.KeyColumnName] = g.Key.keyColumn;
            //            row[_condition.NameColumnWithFilteringService] = g.Key.filteringService;
            //            row[_condition.NameColumnWithFilteringServiceValue] = g.Key.filteringServiceValue;
            //            row["Результат"] = g.Sum(r => r.Value);

            //            // Status?.Invoke(this, new TextEventArgs($"Method: {nameof(MakePivotDataTable1)}"));

            //            return row;
            //        });

            return result.CopyToDataTable();
        }

        public virtual DataTable FilterDataTable()
        {
            SourceDataTableInfo(nameof(FilterDataTable));

            return (from myRow in Source.AsEnumerable()
                    where myRow.Field<String>(_condition.NameColumnWithFilteringService) == _condition.FilteringService
                    select myRow).CopyToDataTable();
        }

        /// <summary>
        /// Used EPPlus
        /// https://stackoverrun.com/ru/q/3109752
        /// </summary>
        /// <param name="path"></param>
        //public void ExportDataTableToPExcelPivot(string path)
        //{
        //    using (DataTable table = Source)
        //    {
        //        System.IO.FileInfo fileInfo = new System.IO.FileInfo(path);
        //        using (var excel = new ExcelPackage(fileInfo))
        //        {
        //            using (var wsData = excel.Workbook.Worksheets.Add("Data-Worksheetname"))
        //            {
        //                wsData.Cells["A1"].LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Medium6);
        //                if (table.Rows.Count != 0)
        //                {
        //                    foreach (DataColumn col in table.Columns)
        //                    {
        //                        // format all dates in german format (adjust accordingly) 
        //                        if (col.DataType == typeof(System.DateTime))
        //                        {
        //                            var colNumber = col.Ordinal + 1;
        //                            var range = wsData.Cells[2, colNumber, table.Rows.Count + 1, colNumber];
        //                            range.Style.Numberformat.Format = "dd.MM.yyyy";
        //                        }
        //                    }
        //                }

        //                using (var dataRange = wsData.Cells[wsData.Dimension.Address.ToString()])
        //                {
        //                    dataRange.AutoFitColumns();

        //                    using (var wsPivot = excel.Workbook.Worksheets.Add("Pivot-Worksheetname"))
        //                    {
        //                    //    var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], dataRange, "Pivotname");
        //                    //    pivotTable.MultipleFieldFilters = true;
        //                    //    pivotTable.RowGrandTotals = true;
        //                    //    pivotTable.ColumGrandTotals = true;
        //                    //    pivotTable.Compact = true;
        //                    //    pivotTable.CompactData = true;
        //                    //    pivotTable.GridDropZones = false;
        //                    //    pivotTable.Outline = false;
        //                    //    pivotTable.OutlineData = false;
        //                    //    pivotTable.ShowError = true;
        //                    //    pivotTable.ErrorCaption = "[error]";
        //                    //    pivotTable.ShowHeaders = true;
        //                    //    pivotTable.UseAutoFormatting = true;
        //                    //    pivotTable.ApplyWidthHeightFormats = true;
        //                    //    pivotTable.ShowDrill = true;
        //                    //    pivotTable.FirstDataCol = 3;
        //                    //    pivotTable.RowHeaderCaption = "Claims";

        //                    //    var modelField = pivotTable.Fields["Model"];
        //                    //    pivotTable.PageFields.Add(modelField);
        //                    //    modelField.Sort = OfficeOpenXml.Table.PivotTable.eSortType.Ascending;

        //                    //    var countField = pivotTable.Fields["Claims"];
        //                    //    pivotTable.DataFields.Add(countField);

        //                    //    var countryField = pivotTable.Fields["Country"];
        //                    //    pivotTable.RowFields.Add(countryField);
        //                    //    var gspField = pivotTable.Fields["GSP/DRSL"];
        //                    //    pivotTable.RowFields.Add(gspField);

        //                    //    var oldStatusField = pivotTable.Fields["Old Status"];
        //                    //    pivotTable.ColumnFields.Add(oldStatusField);
        //                    //    var newStatusField = pivotTable.Fields["New Status"];
        //                    //    pivotTable.ColumnFields.Add(newStatusField);

        //                    //    var submittedDateField = pivotTable.Fields["Claim Submitted Date"];
        //                    //    pivotTable.RowFields.Add(submittedDateField);
        //                    //    submittedDateField.AddDateGrouping(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months | OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
        //                    //    var monthGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Months);
        //                    //    monthGroupField.ShowAll = false;
        //                    //    var dayGroupField = pivotTable.Fields.GetDateGroupField(OfficeOpenXml.Table.PivotTable.eDateGroupBy.Days);
        //                    //    dayGroupField.ShowAll = false;
        //                    }
        //                }
        //            }

        //            excel.Save();
        //        }
        //    }
        //}
    }

    /// <summary>
    /// It will always return a copy of DataTable
    /// </summary>
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
        DataTable MakePivotDataTable2();
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
        /// name of the Column in which are the calculating data
        /// </summary>
        public string NameColumnWithFilteringServiceValue { get; set; }

        /// <summary>
        /// array of Column Names in order for using GroupBy selection
        /// </summary>
        public string[] GroupByOrderColumns { get; set; }
        
      //  public string FilteringServiceValue { get; set; }

        /// <summary>
        /// Type of calculated data
        /// </summary>
     //   public TypeData TypeResultCalcultedData { get; set; }
    }

    //[Flags]
    //public enum TypeData
    //{
    //    None = 0,
    //    DataBool = 1,
    //    DataInt = 2,
    //    DataLong = 4,
    //    DataDouble = 8,
    //    DataString = 16,
    //    DataStringMb = 32,
    //    DataStringkB = 64,
    //    DataStringB = 128,
    //}
}
