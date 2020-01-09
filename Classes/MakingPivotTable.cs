using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    public class MakingPivotDataTable //:  IFilterableDataTable
    {
        public delegate void MessageStatus(object sender, TextEventArgs e);
        public event MessageStatus Status;

        ConditionForMakingPivotTable _condition;

        DataTable _source;

        public MakingPivotDataTable() { }

        public MakingPivotDataTable(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            _source = dataTable;
            _condition = condition;

            PrepareTable(ref _source, ref _condition);
        }

        public void PrepareTable(ref DataTable dataTable, ref ConditionForMakingPivotTable condition)
        {
            DataColumn column = dataTable.Columns.Add(condition.NameColumnWithResult, System.Type.GetType("System.Decimal"));
            foreach (System.Data.DataColumn col in dataTable.Columns)
            { col.ReadOnly = false; }

            //column.Expression = $"{_condition.NameColumnWithFilteringServiceValue } * Quantity";            

            foreach (DataRow row in dataTable.Rows)
            {
                string cell = row[condition.NameColumnWithFilteringService]?.ToString()?.Trim()?.ToUpper();
                if (cell != null && cell.Contains("INTERNET"))
                {
                    row[condition.NameColumnWithResult] = row[condition.NameColumnWithFilteringServiceValue]?.ToString()?.ToInternetTrafic("Mb");
                }
                else
                    row[condition.NameColumnWithResult] = 0;
            }
            dataTable.AcceptChanges();

           string[] orderColumns = EnlargeArray(condition.GroupByOrderColumns, condition.NameColumnWithResult);
            dataTable.SetColumnsOrder(orderColumns);
            condition.GroupByOrderColumns = orderColumns;
          //  _source.AcceptChanges();

            Status?.Invoke(this, new TextEventArgs($"ColumnNames new: {dataTable.ExportColumnInfoToText()}"));
            Status?.Invoke(this, new TextEventArgs($"{condition.GroupByOrderColumns.ToList().AsString(Environment.NewLine).ToString()}"));
        }


        public string[] EnlargeArray(string[] array, string addToList)
        {
            if (addToList == null || addToList.Length == 0)
                return array;

            List<string> columns = array.ToList();
            columns.Add(addToList);
            string[] temp = columns.ToArray();

            return temp;
        }


        public DataTable GroupBy(DataTable source, string groupByColumn)
        {
            DataView dv = new DataView(source);

            //getting distinct values for group column
            DataTable dtGroup = dv.ToTable(true, new string[] { groupByColumn });

            //returning grouped/counted result
            return dtGroup;
        }
        public DataTable ComputeAndGroupBy(string i_sGroupByColumn, string i_sAggregateColumn, DataTable i_dSourceTable)
        {
            DataView dv = new DataView(i_dSourceTable);

            //getting distinct values for group column
            DataTable dtGroup = dv.ToTable(true, new string[] { i_sGroupByColumn });

            //adding column for the row count
            //   dtGroup.Columns.Add("Sum", typeof(decimal));

            //looping thru distinct values for the group, counting
            foreach (DataRow dr in dtGroup.Rows)
            {
                //     dr["Sum"] = i_dSourceTable.Compute("Sum(" + i_sAggregateColumn + ")", i_sGroupByColumn + " = '" + dr[i_sGroupByColumn] + "'");
            }

            //returning grouped/counted result
            return dtGroup;
        }

        private DataColumnCollection SourceDataTableInfo(string calledByMethod)
        {
            Status?.Invoke(this, new TextEventArgs($"Calling method: {calledByMethod}"));
            Status?.Invoke(this, new TextEventArgs($"Колонок в таблице: {_source.Columns.Count}"));

            //List<string> columnsName = new List<string>();
            DataColumnCollection columns = _source.Columns;
            foreach (DataColumn dc in columns)
            {
                //columnsName.Add(dc.ColumnName);
            }
            //Status?.Invoke(this, new TextEventArgs($"Колонки в таблице:{Environment.NewLine}{columnsName.AsString(Environment.NewLine)}"));
            
            Status?.Invoke(this, new TextEventArgs($"Строк в таблице: {_source.Rows.Count}"));

            return columns;
        }


        public virtual DataTable MakePivotDataTable6() //Doesn't correct
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable6));
            var result = _source.AsEnumerable()
              .GroupBy(r => r.Field<int>("Id"))
              .Select(g =>
              {
                  var row = _source.NewRow();

                  row["Id"] = g.Key;
                  row["Amount 1"] = g.Sum(r => r.Field<int>("Amount 1"));
                  row["Amount 2"] = g.Sum(r => r.Field<int>("Amount 2"));
                  row["Amount 3"] = g.Sum(r => r.Field<int>("Amount 3"));

                  return row;
              }).CopyToDataTable();

            return result;
        }

        public virtual IEnumerable MakePivotDataTable5() //Doesn't correct
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable5));
            var result = from row in _source.AsEnumerable()
                         group row by row.Field<string>(_condition.NameColumnWithFilteringServiceValue).ToInternetTrafic("Mb") 
                         into grp
                         select new
                         {
                             TeamID = grp.Key, //not exist
                             MemberCount = grp.Count()
                         }
                         ;

            return result;
        }


        public virtual DataTable MakePivotDataTable4()
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable4));
            DataTable result = _source.AsEnumerable()
                        .GroupBy(r =>  r.Field<string>(_condition.KeyColumnName) )//, Col2 = _condition.GroupByOrderColumns
                        .Select(g => g.OrderBy(x => x.ItemArray = _condition.GroupByOrderColumns).First())
                        .CopyToDataTable();
            
            return result;
        }

        public virtual DataTable MakePivotDataTable3()
        {
          //SourceDataTableInfo(nameof(MakePivotDataTable3));
            DataTable result = GroupBy(_source,_condition.KeyColumnName);
            
            return result;
        }

        public virtual DataTable MakePivotDataTable2()
        {

          return  MakePivotDataTable2(_source);
        }

        public virtual DataTable MakePivotDataTable2(DataTable source)
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable2));

            List<string> removeColumns  = source.ExportColumnNameToList().Except(_condition.GroupByOrderColumns.ToList()).ToList();
            
            //Status?.Invoke(this, new TextEventArgs($"List from DB:{Environment.NewLine}{source.ExportColumnNameToList().AsString(Environment.NewLine)}"));
            //Status?.Invoke(this, new TextEventArgs($"List from condition:{Environment.NewLine}{_condition.GroupByOrderColumns.ToList().AsString(Environment.NewLine)}"));
             
            //Status?.Invoke(this, new TextEventArgs($"List need to remove:{Environment.NewLine}{removeColumns.AsString(Environment.NewLine)}"));
           foreach (var col in removeColumns)
            {
                try { source.RemoveColumn(col); }
                catch (Exception err)
                {
                    Status?.Invoke(this, new TextEventArgs($"{col}\n{err.ToString()}"));
                }
            }
            
            DataTable result = source.AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringService)
                        .Contains(_condition.FilteringService))
                ?.CopyToDataTable();
           
            return result;
        }

        //Do PivotTable
        public virtual DataTable MakePivotDataTable1()
        {
            return MakePivotDataTable1(MakePivotDataTable2());
        }
        public virtual DataTable MakePivotDataTable1(DataTable source)
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable1));
            //var result = source.AsEnumerable()
            //            .GroupBy(r => r[_condition.KeyColumnName])
            //            .Select(g => g)
            //            ;

            var result = source.AsEnumerable()
                .Select(a => new
                {
                    keyColumn = a.Field<string>(_condition.KeyColumnName),
                    filteringService = a.Field<string>(_condition.NameColumnWithFilteringService),
                  //  filteringServiceValue = a.Field<string>(_condition.NameColumnWithResult),
                    Value = a.Field<decimal>(_condition.NameColumnWithResult),
                })
                .GroupBy(r => new { r.keyColumn,  r.filteringService, r.Value })
                .Select(g =>
                    {
                        var row = source.NewRow();

                        row[_condition.KeyColumnName] = g.Key.keyColumn;
                        row[_condition.NameColumnWithFilteringService] = g.Key.filteringService;
                       // row[_condition.NameColumnWithFilteringServiceValue] = g.Key.filteringServiceValue;
                        row[_condition.NameColumnWithResult] = g.Sum(r => r.Value);

                        // Status?.Invoke(this, new TextEventArgs($"Method: {nameof(MakePivotDataTable1)}"));

                        return row;
                    });

            return result.CopyToDataTable();
        }

        public virtual DataTable FilterDataTable(DataTable source)
        {
            //SourceDataTableInfo(nameof(FilterDataTable));
            DataTable result = (from myRow in source.AsEnumerable()
                    where myRow.Field<String>(_condition.NameColumnWithFilteringService) == _condition.FilteringService
                    select myRow).CopyToDataTable();

            return result;
        }
    }

    /// <summary>
    /// It will always return a copy of DataTable
    /// </summary>
    //public abstract class Datatable
    //{
    //    DataTable _source;
    //    public DataTable Source
    //    {
    //        get { return _source.Copy(); }
    //        set { _source = value; }
    //    }
    //}

    //public interface IFilterableDataTable
    //{
    //    DataTable MakePivotDataTable1(DataTable source);
    //    DataTable MakePivotDataTable2(DataTable source);
    //}

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
        
        public string NameColumnWithResult { get; set; }

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
