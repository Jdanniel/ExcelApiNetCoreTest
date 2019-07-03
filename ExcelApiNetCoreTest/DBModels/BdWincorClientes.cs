using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdWincorClientes
    {
        public int IdClienteWincor { get; set; }
        public string Nombre { get; set; }
        public int? IdCliente { get; set; }
    }
}
