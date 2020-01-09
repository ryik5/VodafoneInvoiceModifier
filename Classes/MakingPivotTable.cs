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
            DataColumn column = dataTable.Columns.Add(condition.NameNewColumnWithResult, System.Type.GetType("System.Decimal"));
            foreach (System.Data.DataColumn col in dataTable.Columns)
            { col.ReadOnly = false; }

            //column.Expression = $"{_condition.NameColumnWithFilteringServiceValue } * Quantity";            

            foreach (DataRow row in dataTable.Rows)
            {
                string cell = row[condition.NameColumnWithFilteringService]?.ToString()?.Trim()?.ToUpper();
                if (cell != null && cell.Contains("INTERNET"))
                {
                    row[condition.NameNewColumnWithResult] = row[condition.NameColumnWithFilteringServiceValue]?.ToString()?.ToInternetTrafic("Mb");
                }
                else
                    row[condition.NameNewColumnWithResult] = 0;
            }
            dataTable.AcceptChanges();

           string[] orderColumns = EnlargeArray(condition.GroupByOrderColumns, condition.NameNewColumnWithResult);
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


        //public DataTable GroupBy(DataTable source, string groupByColumn)
        //{
        //    DataView dv = new DataView(source);

        //    //getting distinct values for group column
        //    DataTable dtGroup = dv.ToTable(true, new string[] { groupByColumn });

        //    //returning grouped/counted result
        //    return dtGroup;
        //}
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


        //public virtual DataTable MakePivotDataTable6() //Doesn't correct
        //{
        //    //SourceDataTableInfo(nameof(MakePivotDataTable6));
        //    var result = _source.AsEnumerable()
        //      .GroupBy(r => r.Field<int>("Id"))
        //      .Select(g =>
        //      {
        //          var row = _source.NewRow();

        //          row["Id"] = g.Key;
        //          row["Amount 1"] = g.Sum(r => r.Field<int>("Amount 1"));
        //          row["Amount 2"] = g.Sum(r => r.Field<int>("Amount 2"));
        //          row["Amount 3"] = g.Sum(r => r.Field<int>("Amount 3"));

        //          return row;
        //      }).CopyToDataTable();

        //    return result;
        //}

        //public virtual IEnumerable MakePivotDataTable5() //Doesn't correct
        //{
        //    //SourceDataTableInfo(nameof(MakePivotDataTable5));
        //    var result = from row in _source.AsEnumerable()
        //                 group row by row.Field<string>(_condition.NameColumnWithFilteringServiceValue).ToInternetTrafic("Mb") 
        //                 into grp
        //                 select new
        //                 {
        //                     TeamID = grp.Key, //not exist
        //                     MemberCount = grp.Count()
        //                 }
        //                 ;

        //    return result;
        //}


        //public virtual DataTable MakePivotDataTable4()
        //{
        //    //SourceDataTableInfo(nameof(MakePivotDataTable4));
        //    DataTable result = _source.AsEnumerable()
        //                .GroupBy(r =>  r.Field<string>(_condition.KeyColumnName) )//, Col2 = _condition.GroupByOrderColumns
        //                .Select(g => g.OrderBy(x => x.ItemArray = _condition.GroupByOrderColumns).First())
        //                .CopyToDataTable();
            
        //    return result;
        //}

      //  public virtual DataTable MakePivotDataTable3()
       // {
          //SourceDataTableInfo(nameof(MakePivotDataTable3));
         //   DataTable result = GroupBy(_source,_condition.KeyColumnName);
            
          //  return result;
       // }

        public virtual DataTable MakePivotDataTable2()
        {
            DataTable dt = MakePivotDataTable2(_source);
                    
            var pivotData = dt.AsEnumerable()
                        .Select(r => new
                        {
                            KeyColumnName = r.Field<string>(_condition.KeyColumnName),
                            Filter = r.Field<string>(_condition.NameColumnWithFilteringService),
                            Result = r.Field<decimal>(_condition.NameNewColumnWithResult)
                        })
                            .GroupBy(x => x.KeyColumnName)
                            .Select(g => new
                            {
                                KeyColumnName = g.Key,
                                Filter=g.Select(c=>c.Filter),
                                Result = g.Sum(c => c.Result)
                            });

           DataTable pivotTable = _source.Clone();
            foreach(var v in pivotData)
            {
                DataRow row = pivotTable.NewRow();

                row[_condition.KeyColumnName] = v.KeyColumnName;
                row[_condition.NameColumnWithFilteringService] = v.Filter;
                row[_condition.NameNewColumnWithResult] = v.Result;

                pivotTable.Rows.Add(row);
            }
            
            return pivotTable;
        }
        
        //Do PivotTable
        public virtual DataTable MakePivotDataTable1()
        {
            DataTable dt = MakePivotDataTable2(_source);
            var pivotData = from s in dt.AsEnumerable()
                             group s by new
                             {
                                 KeyColumn = s.Field<string>(_condition.KeyColumnName),
                                 FilterColumn = s.Field<string>(_condition.NameColumnWithFilteringService),
                                 Result = s.Field<decimal>(_condition.NameNewColumnWithResult),
                             } into g
                             select new
                             {
                                 KeyColumn = g.Key.KeyColumn,
                                 FilterColumn = g.Key.FilterColumn,
                                 Result = g.Sum(s => g.Key.Result)
                             };
            
            DataTable pivotTable = _source.Clone();
            foreach (var v in pivotData)
            {
                DataRow row = pivotTable.NewRow();

                row[_condition.KeyColumnName] = v.KeyColumn;
                row[_condition.NameColumnWithFilteringService] = v.FilterColumn;
                row[_condition.NameNewColumnWithResult] = v.Result;

                pivotTable.Rows.Add(row);
            }
            
            return pivotTable;
        }


        /// <summary>
        /// Filter data, Set collection columns and columns' order
        /// </summary>
        /// <param name="source">DataTable with Data</param>
        /// <returns></returns>
        public virtual DataTable MakePivotDataTable2(DataTable source)
        {
            //SourceDataTableInfo(nameof(MakePivotDataTable2));

            DataTable _source = source;
            ChangeColumnsCollection(ref _source);
                       
            DataTable result = _source.AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringService)
                        .Contains(_condition.FilteringService))
                ?.CopyToDataTable();

            return result;
        }

        private void ChangeColumnsCollection(ref DataTable source)
        {
            List<string> removeColumns = ListRemovingColumnsFromDataTable(source, _condition.GroupByOrderColumns);
            foreach (var col in removeColumns)
            {
                try { source.RemoveColumn(col); }
                catch (Exception err)
                {
                    Status?.Invoke(this, new TextEventArgs($"{col}\n{err.ToString()}"));
                }
            }
        }

        /// <summary>
        /// Reorder columns' collection in DataTable
        /// </summary>
        /// <param name="source">DataTable which columns' collection is going to change</param>
        /// <param name="columnsAlive">columns' collection which has to put into or stay in DataTable</param>
        /// <returns></returns>
        private List<string> ListRemovingColumnsFromDataTable(DataTable source, string[] columnsAlive)
        {
            List<string> columns  = source
                .ExportColumnNameToList()
                .Except(columnsAlive.ToList())
                .ToList();
            return columns;
        }


        //public virtual DataTable FilterDataTable(DataTable source)
        //{
        //    //SourceDataTableInfo(nameof(FilterDataTable));
        //    DataTable result = (from myRow in source.AsEnumerable()
        //            where myRow.Field<String>(_condition.NameColumnWithFilteringService) == _condition.FilteringService
        //            select myRow).CopyToDataTable();

        //    return result;
        //}
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
        
        public string NameNewColumnWithResult { get; set; }

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
