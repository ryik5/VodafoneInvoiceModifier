using System.Data;
using System.Linq;

namespace MobileNumbersDetailizationReportGenerator
{

    public class MakerPivotTable
    {
        ConditionForMakingPivotTable _condition;
        DataTable _source;

        public MakerPivotTable(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            _condition = ExpandOrderColumnsInConditionWithNewColumns(condition);

            _source = SetInternetTrafic(
                AddNewColumnsAtDataTable(dataTable, _condition)
                .SetColumnsOrder(condition.ColumnsCollectionAtRightOrder), _condition);
        }

        public DataTable SetInternetTrafic(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            DataTable dt = dataTable.Copy();
            foreach (DataRow row in dt.Rows)
            {
                string cell = row[condition.NameColumnWithFilteringService]?.ToString()?.Trim()?.ToUpper();
                if (cell != null && cell.Contains(condition.FilteringService.ToUpper()))
                {
                    row[condition.NameNewColumnWithSummary] = row[condition.NameColumnWithFilteringServiceValue]?.ToString()?.ToInternetTrafic("Mb");// ?? 0;
                    row[condition.NameNewColumnWithCount] = row[condition.NameColumnWithFilteringServiceValue]?.ToString()?.ToInternetTrafic("Mb") > 0 ? 1 : 0; //только для тех у кого был трафик будет отличный от нуля результат
                }
                // иначе при генерации сводной таблицы в линк-запросе в MakePivot() будет ошибка (не обрабатывает данные)
                // или же предварительно перед MakePivot() выполнять фильтрование записей в Filter() - будут отсутствовать записи с отсутствующим значением, т.е. там где был не трафик, а звонки или смс
                else
                {
                    row[condition.NameNewColumnWithSummary] = 0; 
                    row[condition.NameNewColumnWithCount] = 0;
                }
            }
            dt.AcceptChanges();

            return dt;
        }

        public ConditionForMakingPivotTable ExpandOrderColumnsInConditionWithNewColumns(ConditionForMakingPivotTable condition)
        {
            ConditionForMakingPivotTable result = condition;

            if (condition?.NameNewColumnWithSummary?.Length > 0 && condition?.NameNewColumnWithCount?.Length > 0)
            {
                string[] orderColumns = condition.ColumnsCollectionAtRightOrder
                 .ExpandArray(condition.NameNewColumnWithSummary)
                 .ExpandArray(condition.NameNewColumnWithCount);
                result.ColumnsCollectionAtRightOrder = orderColumns;
            }

            return result;
        }

        private DataTable AddNewColumnsAtDataTable(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            DataTable dt = dataTable.Copy();
            // DataColumn column = 
            dt.Columns.Add(condition.NameNewColumnWithSummary, System.Type.GetType("System.Decimal"));
            dt.Columns.Add(condition.NameNewColumnWithCount, System.Type.GetType("System.Int32"));
            foreach (System.Data.DataColumn col in dt.Columns)
            { col.ReadOnly = false; }

            //column.Expression = $"{_condition.NameColumnWithFilteringServiceValue } * Quantity";            
            dt.AcceptChanges();

            return dt;
        }

        /*
                //public DataTable GroupBy(DataTable source, string groupByColumn)
                //{
                //    DataView dv = new DataView(source);

                //    //getting distinct values for group column
                //    DataTable dtGroup = dv.ToTable(true, new string[] { groupByColumn });

                //    //returning grouped/counted result
                //    return dtGroup;
                //}
                //public DataTable ComputeAndGroupBy(string i_sGroupByColumn, string i_sAggregateColumn, DataTable i_dSourceTable)
                //{
                //    DataView dv = new DataView(i_dSourceTable);

                //    //getting distinct values for group column
                //    DataTable dtGroup = dv.ToTable(true, new string[] { i_sGroupByColumn });

                //    //adding column for the row count
                //    //   dtGroup.Columns.Add("Sum", typeof(decimal));

                //    //looping thru distinct values for the group, counting
                //    foreach (DataRow dr in dtGroup.Rows)
                //    {
                //        //     dr["Sum"] = i_dSourceTable.Compute("Sum(" + i_sAggregateColumn + ")", i_sGroupByColumn + " = '" + dr[i_sGroupByColumn] + "'");
                //    }

                //    //returning grouped/counted result
                //    return dtGroup;
                //}

                //private DataColumnCollection SourceDataTableInfo(string calledByMethod)
                //{
                //    List<string> columnsName = new List<string>();
                //    DataColumnCollection columns = _source.Columns;
                //    foreach (DataColumn dc in columns)
                //    {
                //        columnsName.Add(dc.ColumnName);
                //    }
                //    Status?.Invoke(this, new TextEventArgs($"Calling method: {calledByMethod}"));
                //    Status?.Invoke(this, new TextEventArgs($"Колонок в таблице: {_source.Columns.Count}{Environment.NewLine}Колонки в таблице:{Environment.NewLine}{columnsName.AsString(Environment.NewLine)}"));
                //    Status?.Invoke(this, new TextEventArgs($"Строк в таблице: {_source.Rows.Count}"));

                //    return columns;
                //}
                */

        public virtual DataTable MakePivot()
        {

            // DataTable dt = Filter(_source, _condition);
            DataTable dt = _source.Copy();

            //{ //return if has dbNull
            //    var hasEmpty = dt
            //      .AsEnumerable()
            //      .Any(x => x.HasErrors);

            //    if (hasEmpty) return dt;
            //}
            //{ //return if has dbNull
            //    var hasEmpty = dt
            //      .AsEnumerable()
            //      .Any(x => x.IsNull(_condition.NameNewColumnWithSummary)||x.IsNull(_condition.NameNewColumnWithCount));

            //    if (hasEmpty) return dt;
            //}


            var pivotData = dt.AsEnumerable()
                        .Select(r => new
                        {
                            KeyColumnName = r.Field<string>(_condition.KeyColumnName),
                            FIO = r.Field<string>("ФИО"),
                            NAV = r.Field<string>("NAV"),
                            Department = r.Field<string>("Подразделение"),
                            Service = r.Field<string>("Номер В"),
                            ResultSummary = r.Field<decimal>(_condition.NameNewColumnWithSummary),
                            ResultCount = r.Field<int>(_condition.NameNewColumnWithCount),
                        })
                        .Where(row => row.Service.Contains(_condition.FilteringService))
                        .GroupBy(x => x.KeyColumnName)
                            .Select(g => new
                            {
                                KeyColumnName = g.Key,
                                FIO = g.Select(c => c.FIO).FirstOrDefault(),
                                NAV = g.Select(c => c.NAV).FirstOrDefault(),
                                Department = g.Select(c => c.Department).FirstOrDefault(),
                                Service = g.Select(c => c.Service).FirstOrDefault(),
                                ResultSummary = g.Sum(c => c.ResultSummary),
                                ResultCount = g.Count(c => c.ResultCount > 0),
                            })
                            .OrderBy(x => x.Department)
                            .ThenBy(x => x.FIO);

            DataTable resultPivot = _source.Clone();
            foreach (var v in pivotData)
            {
                DataRow row = resultPivot.NewRow();

                row[_condition.KeyColumnName] = v.KeyColumnName;
                row["ФИО"] = v.FIO;
                row["NAV"] = v.NAV;
                row["Подразделение"] = v.Department;
                row["Номер В"] = v.Service;
                row[_condition.NameNewColumnWithSummary] = v.ResultSummary;
                row[_condition.NameNewColumnWithCount] = v.ResultCount;

                resultPivot.Rows.Add(row);
            }

            return resultPivot;
        }

        /// <summary>
        /// Filter DataTable
        /// </summary>
        /// <param name="source">DataTable with Data</param>
        /// <returns></returns>
        public DataTable Filter(DataTable source, ConditionForMakingPivotTable condition)
        {
            DataTable result = source
                .AsEnumerable()
                .Where(myRow => myRow.Field<string>(condition.NameColumnWithFilteringService)
                        .Contains(condition.FilteringService))
                ?.CopyToDataTable();

            return result;
        }

        public DataTable Source { get { return _source; } }
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
        public string[] ColumnsCollectionAtRightOrder { get; set; }

        public string NameNewColumnWithSummary { get; set; }

        public string NameNewColumnWithCount { get; set; }

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
