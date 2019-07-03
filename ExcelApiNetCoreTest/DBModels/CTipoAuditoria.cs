using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTipoAuditoria
    {
        public CTipoAuditoria()
        {
            BdAuditorias = new HashSet<BdAuditorias>();
        }

        public int IdTipoAuditoria { get; set; }
        public string DescTipoAuditoria { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdAuditorias> BdAuditorias { get; set; }
    }
}
