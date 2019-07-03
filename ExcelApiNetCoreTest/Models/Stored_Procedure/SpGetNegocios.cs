using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelApiNetCoreTest.Models.Stored_Procedure
{
    public class SpGetNegocios
    {
        public int ID_NEGOCIO { get; set; }
        public string DESC_NEGOCIO { get; set; }
        public string DESC_CLIENTE { get; set; }
        public string NO_AFILIACION { get; set; }
        public string DESC_ZONA { get; set; }
        public string DESC_REGION { get; set; }
        public string DIRECCION { get; set; }
        public string STATUS { get; set; }
        public string TELEFONO { get; set; }
        public string COLONIA { get; set; }
        public string POBLACION { get; set; }
        public string CP { get; set; }
        public string DESC_LOCALIDAD { get; set; }
        public string RFC { get; set; }
        public string RAZON_SOCIAL { get; set; }
    }
}
