using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class CTiempos
    {
        public int IdTiempos { get; set; }
        public int? Horas { get; set; }
        public int? Minutos { get; set; }
        public string Descripcion { get; set; }
        public int? Traslado { get; set; }
        public int? Servicio { get; set; }
        public int? MinTotales { get; set; }
    }
}
