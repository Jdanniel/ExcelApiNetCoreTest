using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CProveedoresUsuarios
    {
        public int IdProveedorUsuario { get; set; }
        public string DescProveedorUsuario { get; set; }
        public DateTime? FecAlta { get; set; }
        public string Status { get; set; }
    }
}
