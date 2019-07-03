using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTipoServicioSolicitudAlmacen
    {
        public CTipoServicioSolicitudAlmacen()
        {
            BdSolicitudesAlmacen = new HashSet<BdSolicitudesAlmacen>();
        }

        public int IdTipoServicioSolicitudAlmacen { get; set; }
        public string DescTipoServicioSolicitudAlmacen { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdSolicitudesAlmacen> BdSolicitudesAlmacen { get; set; }
    }
}
