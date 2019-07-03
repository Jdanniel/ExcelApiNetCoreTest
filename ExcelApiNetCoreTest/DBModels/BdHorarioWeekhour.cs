using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdHorarioWeekhour
    {
        public int IdHorarioWeekhour { get; set; }
        public int? IdHorario { get; set; }
        public int? IdWeekhour { get; set; }
        public int? IdUsuarioAlta { get; set; }
        public DateTime? FecAlta { get; set; }

        public virtual CHorarios IdHorarioNavigation { get; set; }
        public virtual CWeekhours IdWeekhourNavigation { get; set; }
    }
}
