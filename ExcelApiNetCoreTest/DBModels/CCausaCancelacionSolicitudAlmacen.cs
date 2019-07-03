using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CCausaCancelacionSolicitudAlmacen
    {
        public int IdCausaCancelacionSolicitudAlmacen { get; set; }
        public string DescCausaCancelacionSolicitudAlmacen { get; set; }
        public int? IdCliente { get; set; }
        public string Status { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public DateTime? FecAlta { get; set; }
    }
}
