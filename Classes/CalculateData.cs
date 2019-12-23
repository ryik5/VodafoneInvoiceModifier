using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
  public interface CalculatePivotData
    {
        SetDataTable(DataTable dataTable) {
    }

   public class CalculateData: CalculatePivotData
    {
        public CalculateData(DataTable dataTable) : base(DataTable dataTable);
    }
}
