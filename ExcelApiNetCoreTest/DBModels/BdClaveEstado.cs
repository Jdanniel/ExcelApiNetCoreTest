using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdClaveEstado
    {
        public int IdClaveEstado { get; set; }
        public int? IdEstado { get; set; }
        public string Clave { get; set; }
        public DateTime? FecAlta { get; set; }
        public int? IdUsuarioAlta { get; set; }

        public virtual SepomexEstados IdEstadoNavigation { get; set; }
    }
}
