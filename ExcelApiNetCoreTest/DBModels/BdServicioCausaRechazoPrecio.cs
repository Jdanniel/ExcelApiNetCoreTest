﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdServicioCausaRechazoPrecio
    {
        public int IdServicioCausaRechazoPrecio { get; set; }
        public int? IdCliente { get; set; }
        public int? IdServicio { get; set; }
        public int? IdEspecificacionCausaRechazo { get; set; }
        public decimal? Precio { get; set; }
    }
}
