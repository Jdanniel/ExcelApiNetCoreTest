﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdTipoPlazaClienteCp
    {
        public int IdTipoPlazaClienteCp { get; set; }
        public int? IdTipoPlazaCliente { get; set; }
        public string Cp { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public DateTime? FecAlta { get; set; }
    }
}
