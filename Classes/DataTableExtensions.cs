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
        public static List<string> DataTableToText(this DataTable dataTable)
        {
            List<string> result=new List<string>();

            foreach (DataRow dr in dataTable.Rows)
            {
                result.Add(string.Join("\t", dr.ItemArray));

                //foreach (DataColumn dc in dataTable.Columns)
                //{
                //    dr[dc.ColumnName].ToString()
                //}
            }

            return result;
        }
    }
}
