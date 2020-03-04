using System;
using System.Collections.Generic;
using System.Data.Entity.Core.EntityClient;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public static class EFConnectionBuilder
    {

        public static EntityConnectionStringBuilder Build(string providerString)
        {
            string providerName = "System.Data.SqlClient";

            // Build the SqlConnection connection string.
            // Initialize the EntityConnectionStringBuilder.
            EntityConnectionStringBuilder entityBuilder =
                new EntityConnectionStringBuilder
                {
                    //Set the provider name.
                    Provider = providerName,

                    // Set the provider-specific connection string.
                    ProviderConnectionString = providerString,

                    // Set the Metadata location.
                    Metadata = @"res://*/Classes.EF.Tfactura.csdl
                                |res://*/Classes.EF.Tfactura.ssdl
                                |res://*/Classes.EF.Tfactura.msl"
                };

            return entityBuilder;
        }
    }
}
