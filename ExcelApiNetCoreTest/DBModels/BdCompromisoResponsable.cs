﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdCompromisoResponsable
    {
        public int IdCompromisoResponsable { get; set; }
        public int? IdCompromiso { get; set; }
        public int? IdResponsable { get; set; }
        public int? IdArea { get; set; }
    }
}
