using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CNombreJunta
    {
        public int IdNombreJunta { get; set; }
        public string NombreJunta { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public DateTime? FecAlta { get; set; }
        public string Status { get; set; }
        public int? NeedArea { get; set; }
    }
}
