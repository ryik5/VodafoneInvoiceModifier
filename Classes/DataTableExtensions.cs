using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
  public static  class DataTableExtensions
    {
        public static List<string> DataTableToText(this DataTable table)
        {
            List<string> result=new List<string>();

            foreach (DataRow dr in table.Rows)
            {
                result.Add(string.Join("\t", dr.ItemArray));
            }

            return result;
        }

        public static string PrintDataTableColumnInfo(this DataTable table)
        {
            string result = string.Empty;

            // Use a DataTable object's DataColumnCollection.
            DataColumnCollection columns = table.Columns;

            // Print the ColumnName and DataType for each column.
            foreach (DataColumn column in columns)
            {
                result += $"Name: {column.ColumnName}\tType: {column.DataType}{Environment.NewLine}";
            }

            return result;
        }    
    }
}
