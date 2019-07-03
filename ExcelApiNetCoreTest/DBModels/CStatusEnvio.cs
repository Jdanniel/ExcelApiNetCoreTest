using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CStatusEnvio
    {
        public CStatusEnvio()
        {
            BdEnvios = new HashSet<BdEnvios>();
        }

        public int IdStatusEnvio { get; set; }
        public string DescStatusEnvio { get; set; }
        public string Status { get; set; }

        public virtual ICollection<BdEnvios> BdEnvios { get; set; }
    }
}
