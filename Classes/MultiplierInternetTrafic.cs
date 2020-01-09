using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

        /// <summary>
        /// Gb,Mb, Kb, b ot GB,MB, KB, B
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
    internal static class MultiplierInternetTrafic
    {
        public static int Multiplier(string dataMesurementTrafic)
        {
            int result;
            switch (dataMesurementTrafic?.ToUpper()?.Trim())
            {
                case ("GB"):
                    result = 1024 * 1024 * 1024;
                    break;
                case ("MB"):
                    result = 1024 * 1024;
                    break;
                case ("KB"):
                    result = 1024;
                    break;
                case ("B"):
                    result = 1;
                    break;
                default:
                    result = -1;
                    break;
            }
            return result;
        }
    }
}
