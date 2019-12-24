using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{


    public class FilteredRowsCollection : AbstractDataTable, IFilterableRowsCollection
    {
        ConditionForMakingPivotTable _condition;

        public FilteredRowsCollection(DataTable dataTable, ConditionForMakingPivotTable condition)
        {
            Source = dataTable;
            _condition = condition;
        }

        public virtual DataTable FilterSource(DataTable collection)
        {
            DataTable result = collection
                .AsEnumerable()
                .Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue)
                .Contains(_condition.FilteringService)).CopyToDataTable();

            return result;
        }

        //DoPivotTable
        public DataTable DoTableUniqKeyRows(DataTable collection)
        {
           // System.Collections.IEnumerable 
           DataTable     result = collection.AsEnumerable()
            //    .SelectMany(row => collection.AsEnumerable().Where(myRow => myRow.Field<string>(_condition.NameColumnWithFilteringServiceValue)
              // .Contains(_condition.FilteringService)))
                .GroupBy(row => row.Field<string>(_condition.KeyColumnName))
                .Select(g =>
                {
                    var row = collection.NewRow();
                    DataColumnCollection col = collection.Columns;

                    foreach (DataColumn kk in col)
                    {
                        row[kk.ColumnName] = g.Key;
                    }

                    row["Id"] = g.Key;
                    row["Amount 1"] = g.Sum(r => r.Field<int>("Amount 1"));
                    row["Amount 2"] = g.Sum(r => r.Field<int>("Amount 2"));
                    row["Amount 3"] = g.Sum(r => r.Field<int>("Amount 3"));

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

    public interface IFilterableRowsCollection
    {
        DataTable FilterSource(DataTable source);
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
