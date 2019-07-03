using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTiempoMedicion
    {
        public CTiempoMedicion()
        {
            BdReglasStatusAr = new HashSet<BdReglasStatusAr>();
        }

        public int IdTiempoMedicion { get; set; }
        public string DescTiempoMedicion { get; set; }
        public string Descripcion { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdReglasStatusAr> BdReglasStatusAr { get; set; }
    }
}
