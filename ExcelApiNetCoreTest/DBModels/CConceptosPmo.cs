﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CConceptosPmo
    {
        public int IdConceptosPmo { get; set; }
        public string DescConcepto { get; set; }
        public int? IdTipoConcepto { get; set; }
        public int? IdAplica { get; set; }
    }
}
