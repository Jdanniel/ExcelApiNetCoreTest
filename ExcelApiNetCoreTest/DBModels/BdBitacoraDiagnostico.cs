﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdBitacoraDiagnostico
    {
        public int IdBitacoraDiagnostico { get; set; }
        public int? IdUnidad { get; set; }
        public int? IdNivelDiagnostico { get; set; }
        public int? IdTecnicoCambio { get; set; }
        public DateTime? Fecha { get; set; }
    }
}
