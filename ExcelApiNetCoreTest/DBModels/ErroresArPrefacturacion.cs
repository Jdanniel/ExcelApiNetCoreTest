﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class ErroresArPrefacturacion
    {
        public int IdArError { get; set; }
        public int? IdArPrefacturacion { get; set; }
        public string Error { get; set; }
    }
}
