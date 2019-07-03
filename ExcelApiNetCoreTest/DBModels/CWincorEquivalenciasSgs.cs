using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CWincorEquivalenciasSgs
    {
        public int IdEquivalencia { get; set; }
        public int IdStatusWincor { get; set; }
        public string DescStatusWincor { get; set; }
        public int IdStatusAr { get; set; }

        public virtual CStatusAr IdStatusArNavigation { get; set; }
    }
}
