using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdOrdenCompraNoFactura
    {
        public int IdOrdenCompraNoFactura { get; set; }
        public int? IdOrdenCompra { get; set; }
        public string NoFactura { get; set; }
        public DateTime? FecFactura { get; set; }
        public int? IdUsuarioAlta { get; set; }
    }
}
