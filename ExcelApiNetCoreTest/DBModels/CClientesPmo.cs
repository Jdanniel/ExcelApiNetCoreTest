﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CClientesPmo
    {
        public int IdClienteProducto { get; set; }
        public int? IdCliente { get; set; }
        public int? IdProducto { get; set; }
        public string Status { get; set; }
    }
}
