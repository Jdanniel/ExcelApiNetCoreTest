using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CProductosNegocios
    {
        public int IdProductoNegocio { get; set; }
        public string DescProductoNegocio { get; set; }
        public string Status { get; set; }
        public int? IdUsuarioAlta { get; set; }
    }
}
