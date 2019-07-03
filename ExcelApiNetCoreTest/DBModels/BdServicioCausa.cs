using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdServicioCausa
    {
        public int IdServicioCausa { get; set; }
        public int IdServicio { get; set; }
        public int IdCausa { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public DateTime? FecAlta { get; set; }

        public virtual CCausas IdCausaNavigation { get; set; }
        public virtual CServicios IdServicioNavigation { get; set; }
    }
}
