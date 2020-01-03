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

        //public virtual DataTable MakePivotDataTable1()
        //{
        //    DataTable result = Source
        //        .AsEnumerable()
        //        .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue))
        //        .Contains(_condition.FilteringService)).CopyToDataTable();

        //    return result;
        //}

        //DoPivotTable
        public DataTable MakePivotDataTable1()
        {
            DataTable result = Source.AsEnumerable()
                .Where(row => row.Field<string>(_condition.NameColumnWithFilteringServiceValue).Contains(_condition.FilteringService))
                .GroupBy(row => row.Field<string>(_condition.KeyColumnName))
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

                            row[$"{_condition.FilteringService}, Sum"] = g.Sum(r => Int32.Parse(r.Field<string>(dc.ColumnName)));

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
