using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdTempServicioCierrePda
    {
        public int IdTempServicioCierrePda { get; set; }
        public int? IdServicio { get; set; }
        public int? IsCierrePda { get; set; }
    }
}
