using System;
using System.Collections.Generic;

namespace ExcelApiNetCoreTest.DBModels
{
    public partial class BdArCierreManual
    {
        public int IdCierreManual { get; set; }
        public int? IdAr { get; set; }
        public string NoSerieInstalacion { get; set; }
        public int? IdConectividadInstalacion { get; set; }
        public int? IdAplicativoInstalacion { get; set; }
        public string VersionInstalacion { get; set; }
        public string ComentarioTerminalInstalacion { get; set; }
        public string ArchivoTerminalInstalacion { get; set; }
        public int? SeDejaMismoSim { get; set; }
        public string SimExistente { get; set; }
        public string SimInstalacion { get; set; }
        public int? IdCarrierInstalacion { get; set; }
        public string ComentarioSimInstalacion { get; set; }
        public string ArchivoSimInstalacion { get; set; }
        public string NoSerieRetiro { get; set; }
        public int? IdMarcaRetiro { get; set; }
        public int? IdModeloRetiro { get; set; }
        public int? IdAplicativoRetiro { get; set; }
        public int? IdConectividadRetiro { get; set; }
        public string VersionRetiro { get; set; }
        public string ComentarioTerminalRetiro { get; set; }
        public string ArchivoTerminalRetiro { get; set; }
        public int? NoSeRetiraSim { get; set; }
        public string SimSinRetirar { get; set; }
        public string SimRetiro { get; set; }
        public int? IdCarrierRet { get; set; }
        public string ComentarioSimRetirado { get; set; }
        public string ArchivoSimRetirado { get; set; }
        public DateTime? FechaCierre { get; set; }
        public string Atiende { get; set; }
        public string OtorganteVobo { get; set; }
        public int? IdTipoAtencion { get; set; }
        public int? IdCriterioCambio { get; set; }
        public int? RollosInstalados { get; set; }
        public int? IdTecnico { get; set; }
        public string Discover { get; set; }
        public string IsAmex { get; set; }
        public string ConclusionesAmex { get; set; }
        public int? IdAmex { get; set; }
        public int? AfiliacionAmex { get; set; }
        public string ComentariosServicio { get; set; }
    }
}
