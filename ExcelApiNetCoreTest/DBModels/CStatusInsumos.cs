using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CStatusInsumos
    {
        public CStatusInsumos()
        {
            BdBitacoraInsumosIdStatusInsumoFinNavigation = new HashSet<BdBitacoraInsumos>();
            BdBitacoraInsumosIdStatusInsumoIniNavigation = new HashSet<BdBitacoraInsumos>();
        }

        public int IdStatusInsumo { get; set; }
        public string DescStatusInsumo { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdBitacoraInsumos> BdBitacoraInsumosIdStatusInsumoFinNavigation { get; set; }
        public virtual ICollection<BdBitacoraInsumos> BdBitacoraInsumosIdStatusInsumoIniNavigation { get; set; }
    }
}
