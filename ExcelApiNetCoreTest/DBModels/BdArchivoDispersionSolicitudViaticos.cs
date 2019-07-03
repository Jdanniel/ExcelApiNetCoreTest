using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdArchivoDispersionSolicitudViaticos
    {
        public int IdArchivoDispersionSolicitudViaticos { get; set; }
        public int IdArchivoDispersion { get; set; }
        public int IdSolicitudViaticos { get; set; }
        public string Status { get; set; }

        public virtual BdArchivosDispersion IdArchivoDispersionNavigation { get; set; }
        public virtual BdSolicitudesViaticos IdSolicitudViaticosNavigation { get; set; }
    }
}
