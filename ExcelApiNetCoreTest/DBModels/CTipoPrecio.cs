using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTipoPrecio
    {
        public CTipoPrecio()
        {
            CServicios = new HashSet<CServicios>();
        }

        public int IdTipoPrecio { get; set; }
        public string DescTipoPrecio { get; set; }
        public string Status { get; set; }

        public virtual ICollection<CServicios> CServicios { get; set; }
    }
}
