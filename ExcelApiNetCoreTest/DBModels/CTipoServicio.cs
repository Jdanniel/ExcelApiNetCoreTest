using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTipoServicio
    {
        public CTipoServicio()
        {
            BdTipoServicioProducto = new HashSet<BdTipoServicioProducto>();
            CServicios = new HashSet<CServicios>();
        }

        public int IdTipoServicio { get; set; }
        public string DescTipoServicio { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdTipoServicioProducto> BdTipoServicioProducto { get; set; }
        public virtual ICollection<CServicios> CServicios { get; set; }
    }
}
