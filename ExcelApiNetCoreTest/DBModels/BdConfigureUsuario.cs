﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdConfigureUsuario
    {
        public int IdConfigureUsuario { get; set; }
        public int? IdUsuario { get; set; }
        public int? Value { get; set; }
        public string Config { get; set; }
        public string Descripcion { get; set; }
    }
}