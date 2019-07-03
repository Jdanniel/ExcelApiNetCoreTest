using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CUrgenciaSolicitudAlmacen
    {
        public CUrgenciaSolicitudAlmacen()
        {
            BdSolicitudesAlmacen = new HashSet<BdSolicitudesAlmacen>();
            BdSolicitudesAlmacenOld = new HashSet<BdSolicitudesAlmacenOld>();
        }

        public int IdUrgenciaSolicitudAlmacen { get; set; }
        public string DescUrgenciaSolicitudAlmacen { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdSolicitudesAlmacen> BdSolicitudesAlmacen { get; set; }
        public virtual ICollection<BdSolicitudesAlmacenOld> BdSolicitudesAlmacenOld { get; set; }
    }
}
