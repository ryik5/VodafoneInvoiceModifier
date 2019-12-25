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

        public MakingPivotDataTable(DataTable dataTable)
        {
            SetDataTable(dataTable);
        }

        public MakingPivotDataTable(ConditionForMakingPivotTable condition)
        {
            SetFilter(condition);
        }

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

        public virtual DataTable MakePivotDataTable1()
        {
            DataTable result = Source
                .AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue)
                .Contains(_condition.FilteringService)).CopyToDataTable();

            return result;
        }

        //DoPivotTable
        public DataTable MakePivotDataTable2()
        {
            DataTable result = Source.AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue) == _condition.FilteringService)
                .GroupBy(row => row.Field<string>(_condition.KeyColumnName))
                .Select(g =>
                {
                    var row = Source.NewRow();
                    DataColumnCollection col = Source.Columns;

                    foreach (DataColumn dc in col)
                    {
                        if (_condition.KeyColumnName.Equals(dc.ColumnName))
                        { row[dc.ColumnName] = g.Key; }
                        else if (_condition.NameColumnWithFilteringServiceValue.Equals(dc.ColumnName))
                        {
                            row[_condition.FilteringService] = g.Sum(r => Int32.Parse(r.Field<string>(dc.ColumnName)));//???
                        }
                        else
                        {
                            row[dc.ColumnName] = g.Key;
                        }
                    }

                    //  row["Id"] = g.Key;
                    //  row["Amount 1"] = g.Sum(r => r.Field<int>("Amount 1"));
                    //  row["Amount 2"] = g.Sum(r => r.Field<int>("Amount 2"));
                    //  row["Amount 3"] = g.Sum(r => r.Field<int>("Amount 3"));

                    return row;
                }).CopyToDataTable();

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
        public string KeyColumnName { get; set; }

        public string FilteringService { get; set; }
        public string NameColumnWithFilteringService { get; set; }

        public string FilteringServiceValue { get; set; }
        public string NameColumnWithFilteringServiceValue { get; set; }

        public TypeData TypeResultCalcultedData { get; set; }
    }

    [Flags]
    public enum TypeData
    {
        None = 0,
        DataBool = 1,
        DataInt = 2,
        DataLong = 3,
        DataDouble = 4,
        DataString = 8
    }
}
