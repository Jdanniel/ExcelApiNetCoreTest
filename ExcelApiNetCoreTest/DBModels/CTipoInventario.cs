using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTipoInventario
    {
        public CTipoInventario()
        {
            BdAuditoriaReporte = new HashSet<BdAuditoriaReporte>();
        }

        public int IdTipoInventario { get; set; }
        public string DescTipoInventario { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdAuditoriaReporte> BdAuditoriaReporte { get; set; }
    }
}
