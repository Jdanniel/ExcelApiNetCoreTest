using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CPuestos
    {
        public int IdPuesto { get; set; }
        public string DescPuesto { get; set; }
        public string Status { get; set; }
        public DateTime? FechaAlta { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public int? IsResponsable { get; set; }
    }
}
