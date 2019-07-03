﻿using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class NegociosHsbcFujitsu
    {
        public int Id { get; set; }
        public int? IdCliente { get; set; }
        public int? IdRegion { get; set; }
        public int? IdZona { get; set; }
        public int? IdSegmento { get; set; }
        public string NoAfiliacion { get; set; }
        public string DescNegocio { get; set; }
        public string Direccion { get; set; }
        public string Colonia { get; set; }
        public string Poblacion { get; set; }
        public string Estado { get; set; }
        public int? IdEstado { get; set; }
        public string Cp { get; set; }
        public string Telefono { get; set; }
        public int? IdHorarioAcceso { get; set; }
        public int? IdHorarioUptime { get; set; }
        public int? IsDuplicated { get; set; }
        public int? IsSystemDuplicated { get; set; }
        public int? IsProcesada { get; set; }
    }
}
