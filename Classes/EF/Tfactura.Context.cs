﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace MobileNumbersDetailizationReportGenerator
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class EBPEntities : DbContext
    {
        public EBPEntities(string connection)
            : base(connection)
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<os_contract_link> os_contract_link { get; set; }
        public virtual DbSet<v_dp_contract_bill_detail_ex> v_dp_contract_bill_detail_ex { get; set; }
        public virtual DbSet<v_rs_contract_detail> v_rs_contract_detail { get; set; }
    }
}
