using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CUnidadMensajeria
    {
        public CUnidadMensajeria()
        {
            BdSolicitudRecoleccion = new HashSet<BdSolicitudRecoleccion>();
        }

        public int IdUnidadMensajeria { get; set; }
        public int IdServicioMensajeria { get; set; }
        public string DescUnidadMensajeria { get; set; }
        public string Status { get; set; }
        public int IdUsuarioAlta { get; set; }
        public DateTime FecAlta { get; set; }

        public virtual ICollection<BdSolicitudRecoleccion> BdSolicitudRecoleccion { get; set; }
    }
}
