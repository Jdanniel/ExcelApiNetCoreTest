using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelApiNetCoreTest.DBModels;
using ExcelApiNetCoreTest.Models;
using ExcelApiNetCoreTest.Models.Stored_Procedure;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelApiNetCoreTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportExcelController : ControllerBase
    {
        protected ELAVONTESTContext context_ = new ELAVONTESTContext();
        private readonly static Random random = new Random();
        private readonly IConfiguration configuration;
        private readonly string appname;

        public ExportExcelController(IConfiguration iConfig)
        {
            configuration = iConfig;
            appname = configuration.GetValue<string>("MySettings:appname");
        }

        [HttpPost("LAYOUT_MASIVO")]
        public async Task<ActionResult> LayoutMasivo(LayoutMasivoRequest layout, CancellationToken cancellationToken)
        {
            if (layout == null)
            {
                return BadRequest("La solicitud esta vacia.");
            }

            try
            {
                string folder = "C:\\inetpub\\wwwroot\\" + appname + "\\REPORTES\\LAYOUT_MASIVO\\ARCHIVOS";
                string name = "Layout_Masivo_OTD.xlsx";
                string downloadUrl = "C:\\inetpub\\wwwroot\\" + appname + "\\REPORTES\\LAYOUT_MASIVO\\ARCHIVOS\\" + name;
                Nullable<Int64> inull = null;
                FileInfo file = new FileInfo(Path.Combine(folder, name));

                var columHeaders = new string[]
                {
                    "ODT",
                    "Discover",
                    "Afiliación",
                    "Comercio",
                    "Dirección",
                    "Colonia",
                    "Ciudad",
                    "Estado",
                    "Fecha Alta",
                    "Fecha Vencimiento",
                    "Descripción",
                    "Observaciones",
                    "Telefono",
                    "Tipo Comercio",
                    "Nivel",
                    "Tipo Servicio",
                    "Subtipo Servicio",
                    "Criterio de Cambio",
                    "Tecnico",
                    "Proveedor",
                    "Estatus Servicio",
                    "Fecha Atención Proveedor",
                    "Fecha Cierre Sistema",
                    "Fecha Alta Sistema",
                    "Codigo Postal",
                    "Conclusiones",
                    "Conectividad",
                    "Modelo",
                    "Id Equipo",
                    "Id Caja",
                    "RFC",
                    "Razón Social",
                    "Horas Vencidas",
                    "Tiempo en atender",
                    "SLA Fijo",
                    "Nivel",
                    "Telefonos en Campo",
                    "Tipo Comercio",
                    "Afiliacion Amex",
                    "IdAmex",
                    "Producto",
                    "Motivo Cancelación",
                    "Motivo Rechazo",
                    "Email",
                    "Rollos a instalar",
                    "Num Serie Terminal Entra",
                    "Num Serie Terminal Sale",
                    "Num Serie Terminal mto",
                    "Num Serie Sim Sale",
                    "Num Serie Sim Entra",
                    "VersionSW",
                    "Cargador",
                    "Base",
                    "Rollo Entregados",
                    "Cable corriente",
                    "Zona",
                    "Modelo Instalado",
                    "Modelo Terminal Sale",
                    "Correo Ejecutivo",
                    "Rechazo",
                    "Contacto 1",
                    "Atiende en comercio",
                    "Tid Amex Cierre",
                    "Afiliacion Amex Cierre",
                    "Codigo",
                    "Tiene Amex",
                    "Act Referencias",
                    "Tipo_A_b",
                    "Domicilio Alterno",
                    "Cantidad Archivos",
                    "Area Carga",
                    "Alta Por",
                    "Tipo Carga",
                    "Cerrado Por",
                    "Serie"
                };
                if (file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(Path.Combine(folder, name));
                }
                List<SpGetLayoutMasivo> procedure = new List<SpGetLayoutMasivo>();

                procedure = await context_.Query<SpGetLayoutMasivo>().FromSql("EXEC SP_LAYOUT_MASIVO @p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8",
                    layout.fec_ini,
                    layout.fec_fin,
                    layout.id_proveedor,
                    layout.status_servicio,
                    layout.id_zona,
                    layout.id_proyecto,
                    layout.fec_ini_cierre,
                    layout.fec_fin_cierre,
                    layout.serie).ToListAsync();

                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Consulta");
                    using (var cells = worksheet.Cells[1, 1, 1, 75])
                    {
                        cells.Style.Font.Bold = true;
                        cells.Style.Font.Color.SetColor(Color.White);
                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 128, 255));
                    }

                    for (int i = 0; i < columHeaders.Count(); i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columHeaders[i];
                    }

                    var j = 2;

                    foreach (var servicio in procedure)
                    {
                        worksheet.Cells["A" + j].Value = servicio.ODT;
                        worksheet.Cells["B" + j].Value = servicio.DISCOVER;
                        worksheet.Cells["C" + j].Style.Numberformat.Format = "0";
                        worksheet.Cells["C" + j].Value = servicio.AFILIACION == "" ? inull : Convert.ToInt64(servicio.AFILIACION);
                        worksheet.Cells["D" + j].Value = servicio.COMERCIO;
                        worksheet.Cells["E" + j].Value = servicio.DIRECCION;
                        worksheet.Cells["F" + j].Value = servicio.COLONIA;
                        worksheet.Cells["G" + j].Value = servicio.POBLACION;
                        worksheet.Cells["H" + j].Value = servicio.ESTADO;
                        worksheet.Cells["I" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["I" + j].Value = servicio.FECHA_ALTA;
                        worksheet.Cells["J" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["J" + j].Value = servicio.FECHA_VENCIMIENTO;
                        worksheet.Cells["K" + j].Value = servicio.DESCRIPCION;
                        worksheet.Cells["L" + j].Value = servicio.OBSERVACIONES;
                        worksheet.Cells["M" + j].Value = servicio.TELEFONO;
                        worksheet.Cells["N" + j].Value = servicio.TIPO_COMERCIO;
                        worksheet.Cells["O" + j].Value = servicio.NIVEL;
                        worksheet.Cells["P" + j].Value = servicio.TIPO_SERVICIO;
                        worksheet.Cells["Q" + j].Value = servicio.SUB_TIPO_SERVICIO;
                        worksheet.Cells["R" + j].Value = servicio.CRITERIO_CAMBIO;
                        worksheet.Cells["S" + j].Value = servicio.ID_TECNICO;
                        worksheet.Cells["T" + j].Value = servicio.PROVEEDOR;
                        worksheet.Cells["U" + j].Value = servicio.ESTATUS_SERVICIO;
                        worksheet.Cells["V" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["V" + j].Value = servicio.FECHA_ATENCION_PROVEEDOR;
                        worksheet.Cells["W" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["W" + j].Value = servicio.FECHA_CIERRE_SISTEMA;
                        worksheet.Cells["X" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["X" + j].Value = servicio.FECHA_ALTA_SISTEMA;
                        worksheet.Cells["Y" + j].Value = servicio.CODIGO_POSTAL;
                        worksheet.Cells["Z" + j].Value = servicio.CONCLUSIONES;
                        worksheet.Cells["AA" + j].Value = servicio.CONECTIVIDAD;
                        worksheet.Cells["AB" + j].Value = servicio.MODELO;
                        worksheet.Cells["AC" + j].Value = servicio.EQUIPO;
                        worksheet.Cells["AD" + j].Value = servicio.CAJA;
                        worksheet.Cells["AE" + j].Value = servicio.RFC;
                        worksheet.Cells["AF" + j].Value = servicio.RAZON_SOCIAL;
                        worksheet.Cells["AG" + j].Value = servicio.HORAS_VENCIDAS;
                        worksheet.Cells["AH" + j].Value = servicio.TIEMPO_EN_ATENDER;
                        worksheet.Cells["AI" + j].Value = servicio.SLA_FIJO;
                        worksheet.Cells["AJ" + j].Value = servicio.NIVEL;
                        worksheet.Cells["AK" + j].Value = servicio.TELEFONOS_EN_CAMPO;
                        worksheet.Cells["AL" + j].Value = servicio.TIPO_COMERCIO;
                        worksheet.Cells["AM" + j].Value = servicio.AFILIACION_AMEX;
                        worksheet.Cells["AN" + j].Value = servicio.IDAMEX;
                        worksheet.Cells["AO" + j].Value = servicio.PRODUCTO;
                        worksheet.Cells["AP" + j].Value = servicio.MOTIVO_CANCELACION;
                        worksheet.Cells["AQ" + j].Value = servicio.MOTIVO_RECHAZO;
                        worksheet.Cells["AR" + j].Value = servicio.EMAIL;
                        worksheet.Cells["AS" + j].Value = servicio.ROLLOS_A_INSTALAR;
                        worksheet.Cells["AT" + j].Value = servicio.NUM_SERIE_TERMINAL_ENTRA;
                        worksheet.Cells["AU" + j].Value = servicio.NUM_SERIE_TERMINAL_SALE;
                        worksheet.Cells["AV" + j].Value = servicio.NUM_SERIE_TERMINAL_MTO;
                        worksheet.Cells["AW" + j].Value = servicio.NUM_SERIE_SIM_SALE;
                        worksheet.Cells["AX" + j].Value = servicio.NUM_SERIE_SIM_ENTRA;
                        worksheet.Cells["AY" + j].Value = servicio.VERSIONSW;
                        worksheet.Cells["AZ" + j].Value = servicio.CARGADOR;
                        worksheet.Cells["BA" + j].Value = servicio.BASE;
                        worksheet.Cells["BB" + j].Value = servicio.ROLLO_ENTREGADOS;
                        worksheet.Cells["BC" + j].Value = servicio.CABLE_CORRIENTE;
                        worksheet.Cells["BD" + j].Value = servicio.ZONA;
                        worksheet.Cells["BE" + j].Value = servicio.MODELO_INSTALADO;
                        worksheet.Cells["BF" + j].Value = servicio.MODELO_TERMINAL_SALE;
                        worksheet.Cells["BG" + j].Value = servicio.CORREO_EJECUTIVO;
                        worksheet.Cells["BH" + j].Value = servicio.RECHAZO;
                        worksheet.Cells["BI" + j].Value = servicio.CONTACTO1;
                        worksheet.Cells["BJ" + j].Value = servicio.ATIENDE_EN_COMERCIO;
                        worksheet.Cells["BK" + j].Value = servicio.TID_AMEX_CIERRE;
                        worksheet.Cells["BL" + j].Value = servicio.AFILIACION_AMEX_CIERRE;
                        worksheet.Cells["BM" + j].Value = servicio.CODIGO;
                        worksheet.Cells["BN" + j].Value = servicio.TIENE_AMEX;
                        worksheet.Cells["BO" + j].Value = servicio.ACT_REFERENCIAS;
                        worksheet.Cells["BP" + j].Value = servicio.TIPO_A_B;
                        worksheet.Cells["BQ" + j].Value = servicio.DIRECCION_ALTERNA_COMERCIO;
                        worksheet.Cells["BR" + j].Value = servicio.CANTIDAD_ARCHIVOS;
                        worksheet.Cells["BS" + j].Value = servicio.AREA_CARGA;
                        worksheet.Cells["BT" + j].Value = servicio.ALTA_POR;
                        worksheet.Cells["BU" + j].Value = servicio.TIPO_CARGA;
                        worksheet.Cells["BV" + j].Value = servicio.CERRADO_POR;
                        worksheet.Cells["BW" + j].Value = servicio.SERIE;

                        j++;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
                return Ok(name);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }
        [HttpPost("LAYOUT_MASIVO2")]
        public async Task<ActionResult> LayoutMasivo2(LayoutMasivoRequest layout, CancellationToken cancellationToken)
        {
            if (layout == null)
            {
                return BadRequest("La solicitud esta vacia.");
            }

            try
            {
                string folder = "C:\\inetpub\\wwwroot\\" + appname + "\\REPORTES\\LAYOUT_MASIVO\\ARCHIVOS";
                string name = "serviciosPorDiaNoProyecto.xlsx";
                string downloadUrl = "C:\\inetpub\\wwwroot\\" + appname + "\\REPORTES\\LAYOUT_MASIVO\\ARCHIVOS\\" + name;
                Nullable<Int64> inull = null;
                FileInfo file = new FileInfo(Path.Combine(folder, name));

                var columHeaders = new string[]
                {
                    "ODT",
                    "Discover",
                    "Afiliación",
                    "Comercio",
                    "Dirección",
                    "Colonia",
                    "Ciudad",
                    "Estado",
                    "Fecha Alta",
                    "Fecha Vencimiento",
                    "Descripción",
                    "Observaciones",
                    "Telefono",
                    "Tipo Comercio",
                    "Nivel",
                    "Tipo Servicio",
                    "Subtipo Servicio",
                    "Criterio de Cambio",
                    "Tecnico",
                    "Proveedor",
                    "Estatus Servicio",
                    "Fecha Atención Proveedor",
                    "Fecha Cierre Sistema",
                    "Fecha Alta Sistema",
                    "Codigo Postal",
                    "Conclusiones",
                    "Conectividad",
                    "Modelo",
                    "Id Equipo",
                    "Id Caja",
                    "RFC",
                    "Razón Social",
                    "Horas Vencidas",
                    "Tiempo en atender",
                    "SLA Fijo",
                    "Nivel",
                    "Telefonos en Campo",
                    "Tipo Comercio",
                    "Afiliacion Amex",
                    "IdAmex",
                    "Producto",
                    "Motivo Cancelación",
                    "Motivo Rechazo",
                    "Email",
                    "Rollos a instalar",
                    "Num Serie Terminal Entra",
                    "Num Serie Terminal Sale",
                    "Num Serie Terminal mto",
                    "Num Serie Sim Sale",
                    "Num Serie Sim Entra",
                    "VersionSW",
                    "Cargador",
                    "Base",
                    "Rollo Entregados",
                    "Cable corriente",
                    "Zona",
                    "Modelo Instalado",
                    "Modelo Terminal Sale",
                    "Correo Ejecutivo",
                    "Rechazo",
                    "Contacto 1",
                    "Atiende en comercio",
                    "Tid Amex Cierre",
                    "Afiliacion Amex Cierre",
                    "Codigo",
                    "Tiene Amex",
                    "Act Referencias",
                    "Tipo_A_b",
                    "Domicilio Alterno",
                    "Cantidad Archivos",
                    "Area Carga",
                    "Alta Por",
                    "Tipo Carga",
                    "Cerrado Por",
                    "Serie"
                };
                if (file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(Path.Combine(folder, name));
                }
                List<SpGetLayoutMasivo> procedure = new List<SpGetLayoutMasivo>();

                procedure = await context_.Query<SpGetLayoutMasivo>().FromSql("EXEC SP_LAYOUT_MASIVO @p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8",
                    layout.fec_ini,
                    layout.fec_fin,
                    layout.id_proveedor,
                    layout.status_servicio,
                    layout.id_zona,
                    layout.id_proyecto,
                    layout.fec_ini_cierre,
                    layout.fec_fin_cierre,
                    layout.serie).ToListAsync();

                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Reporte Servicios Por Dia");
                    using (var cells = worksheet.Cells[1, 1, 1, 75])
                    {
                        cells.Style.Font.Bold = true;
                        cells.Style.Font.Color.SetColor(Color.White);
                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 128, 255));
                    }

                    for (int i = 0; i < columHeaders.Count(); i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columHeaders[i];
                    }

                    var j = 2;

                    foreach (var servicio in procedure)
                    {
                        worksheet.Cells["A" + j].Value = servicio.ODT;
                        worksheet.Cells["B" + j].Value = servicio.DISCOVER;
                        worksheet.Cells["C" + j].Style.Numberformat.Format = "0";
                        worksheet.Cells["C" + j].Value = servicio.AFILIACION == "" ? inull : Convert.ToInt64(servicio.AFILIACION);
                        worksheet.Cells["D" + j].Value = servicio.COMERCIO;
                        worksheet.Cells["E" + j].Value = servicio.DIRECCION;
                        worksheet.Cells["F" + j].Value = servicio.COLONIA;
                        worksheet.Cells["G" + j].Value = servicio.POBLACION;
                        worksheet.Cells["H" + j].Value = servicio.ESTADO;
                        worksheet.Cells["I" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["I" + j].Value = servicio.FECHA_ALTA;
                        worksheet.Cells["J" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["J" + j].Value = servicio.FECHA_VENCIMIENTO;
                        worksheet.Cells["K" + j].Value = servicio.DESCRIPCION;
                        worksheet.Cells["L" + j].Value = servicio.OBSERVACIONES;
                        worksheet.Cells["M" + j].Value = servicio.TELEFONO;
                        worksheet.Cells["N" + j].Value = servicio.TIPO_COMERCIO;
                        worksheet.Cells["O" + j].Value = servicio.NIVEL;
                        worksheet.Cells["P" + j].Value = servicio.TIPO_SERVICIO;
                        worksheet.Cells["Q" + j].Value = servicio.SUB_TIPO_SERVICIO;
                        worksheet.Cells["R" + j].Value = servicio.CRITERIO_CAMBIO;
                        worksheet.Cells["S" + j].Value = servicio.ID_TECNICO;
                        worksheet.Cells["T" + j].Value = servicio.PROVEEDOR;
                        worksheet.Cells["U" + j].Value = servicio.ESTATUS_SERVICIO;
                        worksheet.Cells["V" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["V" + j].Value = servicio.FECHA_ATENCION_PROVEEDOR;
                        worksheet.Cells["W" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["W" + j].Value = servicio.FECHA_CIERRE_SISTEMA;
                        worksheet.Cells["X" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["X" + j].Value = servicio.FECHA_ALTA_SISTEMA;
                        worksheet.Cells["Y" + j].Value = servicio.CODIGO_POSTAL;
                        worksheet.Cells["Z" + j].Value = servicio.CONCLUSIONES;
                        worksheet.Cells["AA" + j].Value = servicio.CONECTIVIDAD;
                        worksheet.Cells["AB" + j].Value = servicio.MODELO;
                        worksheet.Cells["AC" + j].Value = servicio.EQUIPO;
                        worksheet.Cells["AD" + j].Value = servicio.CAJA;
                        worksheet.Cells["AE" + j].Value = servicio.RFC;
                        worksheet.Cells["AF" + j].Value = servicio.RAZON_SOCIAL;
                        worksheet.Cells["AG" + j].Value = servicio.HORAS_VENCIDAS;
                        worksheet.Cells["AH" + j].Value = servicio.TIEMPO_EN_ATENDER;
                        worksheet.Cells["AI" + j].Value = servicio.SLA_FIJO;
                        worksheet.Cells["AJ" + j].Value = servicio.NIVEL;
                        worksheet.Cells["AK" + j].Value = servicio.TELEFONOS_EN_CAMPO;
                        worksheet.Cells["AL" + j].Value = servicio.TIPO_COMERCIO;
                        worksheet.Cells["AM" + j].Value = servicio.AFILIACION_AMEX;
                        worksheet.Cells["AN" + j].Value = servicio.IDAMEX;
                        worksheet.Cells["AO" + j].Value = servicio.PRODUCTO;
                        worksheet.Cells["AP" + j].Value = servicio.MOTIVO_CANCELACION;
                        worksheet.Cells["AQ" + j].Value = servicio.MOTIVO_RECHAZO;
                        worksheet.Cells["AR" + j].Value = servicio.EMAIL;
                        worksheet.Cells["AS" + j].Value = servicio.ROLLOS_A_INSTALAR;
                        worksheet.Cells["AT" + j].Value = servicio.NUM_SERIE_TERMINAL_ENTRA;
                        worksheet.Cells["AU" + j].Value = servicio.NUM_SERIE_TERMINAL_SALE;
                        worksheet.Cells["AV" + j].Value = servicio.NUM_SERIE_TERMINAL_MTO;
                        worksheet.Cells["AW" + j].Value = servicio.NUM_SERIE_SIM_SALE;
                        worksheet.Cells["AX" + j].Value = servicio.NUM_SERIE_SIM_ENTRA;
                        worksheet.Cells["AY" + j].Value = servicio.VERSIONSW;
                        worksheet.Cells["AZ" + j].Value = servicio.CARGADOR;
                        worksheet.Cells["BA" + j].Value = servicio.BASE;
                        worksheet.Cells["BB" + j].Value = servicio.ROLLO_ENTREGADOS;
                        worksheet.Cells["BC" + j].Value = servicio.CABLE_CORRIENTE;
                        worksheet.Cells["BD" + j].Value = servicio.ZONA;
                        worksheet.Cells["BE" + j].Value = servicio.MODELO_INSTALADO;
                        worksheet.Cells["BF" + j].Value = servicio.MODELO_TERMINAL_SALE;
                        worksheet.Cells["BG" + j].Value = servicio.CORREO_EJECUTIVO;
                        worksheet.Cells["BH" + j].Value = servicio.RECHAZO;
                        worksheet.Cells["BI" + j].Value = servicio.CONTACTO1;
                        worksheet.Cells["BJ" + j].Value = servicio.ATIENDE_EN_COMERCIO;
                        worksheet.Cells["BK" + j].Value = servicio.TID_AMEX_CIERRE;
                        worksheet.Cells["BL" + j].Value = servicio.AFILIACION_AMEX_CIERRE;
                        worksheet.Cells["BM" + j].Value = servicio.CODIGO;
                        worksheet.Cells["BN" + j].Value = servicio.TIENE_AMEX;
                        worksheet.Cells["BO" + j].Value = servicio.ACT_REFERENCIAS;
                        worksheet.Cells["BP" + j].Value = servicio.TIPO_A_B;
                        worksheet.Cells["BQ" + j].Value = servicio.DIRECCION_ALTERNA_COMERCIO;
                        worksheet.Cells["BR" + j].Value = servicio.CANTIDAD_ARCHIVOS;
                        worksheet.Cells["BS" + j].Value = servicio.AREA_CARGA;
                        worksheet.Cells["BT" + j].Value = servicio.ALTA_POR;
                        worksheet.Cells["BU" + j].Value = servicio.TIPO_CARGA;
                        worksheet.Cells["BV" + j].Value = servicio.CERRADO_POR;
                        worksheet.Cells["BW" + j].Value = servicio.SERIE;

                        j++;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
                return Ok(name);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }
        [HttpPost("CONSULTA_UNIDADES")]
        public async Task<ActionResult> Unidades(ConsultaUnidadesRequest consulta, CancellationToken cancellationToken)
        {
            if (consulta == null)
            {
                return BadRequest("La solicitud esta vacia.");
            }
            try
            {
                string folder = "C:\\inetpub\\wwwroot\\" + appname + "\\UNIDADES\\ARCHIVOS";
                string name = "Unidades_" + RandomString(10) + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + ".xlsx";
                string downloadUrl = "C:\\inetpub\\wwwroot\\" + appname + "\\UNIDADES\\ARCHIVOS\\" + name;
                Nullable<Int64> inull = null;
                FileInfo file = new FileInfo(Path.Combine(folder, name));

                var columHeaders = new string[]
                {
                    "REPORTE",
                    "PROVEEDOR",
                    "AFILIACION",
                    "SERIE",
                    "CONTACLESS SI/NO",
                    "NUMERO DE VECES QUE ESTUVO EN REPARACION",
                    "ESTADO",
                    "ZONA",
                    "CODIGO POSTAL",
                    "ULTIMA DIRECCION",
                    "APLICATIVO",
                    "CONECTIVIDAD",
                    "PRODUCTO",
                    "CATEGORIA",
                    "PRECIO INICIO",
                    "MODELO",
                    "FACTURA_NO",
                    "FECHA_FACTURA",
                    "COSTO CON LA DEPRECIACION",
                    "ESTATUS",
                    "FECHA_ASIGNACION",
                    "OBSERVACION/MOTIVO DE ENAJENACION",
                    "ODT",
                    "FECHA ATENCION PROVEEDOR"
                };

                if (file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(Path.Combine(folder, name));
                }
                List<SpGetConsultaUnidades> procedure = new List<SpGetConsultaUnidades>();

                procedure = await context_.Query<SpGetConsultaUnidades>().FromSql("EXEC SP_GET_CONSULTA_UNIDADES " +
                    "@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12",
                        consulta.search_text,
                        consulta.desc_unidad,
                        consulta.idcategoria,
                        consulta.idproveedor,
                        consulta.isdaniada,
                        consulta.idresponsable,
                        consulta.idtiporesponsable,
                        consulta.idaplicativo,
                        consulta.idconectividad,
                        consulta.idcliente,
                        consulta.idproducto,
                        consulta.isnueva,
                        consulta.idusuario
                    ).ToListAsync();

                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Unidades");
                    using (var cells = worksheet.Cells[1, 1, 1, 24])
                    {
                        cells.Style.Font.Bold = true;
                        cells.Style.Font.Color.SetColor(Color.White);
                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 128, 255));
                    }

                    for (int i = 0; i < columHeaders.Count(); i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columHeaders[i];
                    }

                    var j = 2;
                    foreach (var unidades in procedure)
                    {
                        worksheet.Cells["A" + j].Value = unidades.REPORTE;
                        worksheet.Cells["B" + j].Value = unidades.PROVEEDOR_2;
                        worksheet.Cells["C" + j].Style.Numberformat.Format = "0";
                        worksheet.Cells["C" + j].Value = unidades.NO_AFILIACION == "" ? inull : Convert.ToInt64(unidades.NO_AFILIACION);
                        worksheet.Cells["D" + j].Style.Numberformat.Format = "0";
                        worksheet.Cells["D" + j].Value = unidades.NO_SERIE;
                        worksheet.Cells["E" + j].Value = unidades.CONTACLESS;
                        worksheet.Cells["F" + j].Value = unidades.NO_REPARACION;
                        worksheet.Cells["G" + j].Value = unidades.ESTADO_RESPONSABLE;
                        worksheet.Cells["H" + j].Value = unidades.ZONA_RESPONSABLE;
                        worksheet.Cells["I" + j].Value = unidades.CP;
                        worksheet.Cells["J" + j].Value = unidades.DIRECCION;
                        worksheet.Cells["K" + j].Value = unidades.APLICATIVO;
                        worksheet.Cells["L" + j].Value = unidades.CONECTIVIDAD;
                        worksheet.Cells["M" + j].Value = unidades.PRODUCTO;
                        worksheet.Cells["N" + j].Value = unidades.CATEGORIA_2;
                        worksheet.Cells["O" + j].Value = unidades.PRECIO_INICIAL;
                        worksheet.Cells["P" + j].Value = unidades.MODELO;
                        worksheet.Cells["Q" + j].Value = unidades.NO_FACTURA;
                        worksheet.Cells["R" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["R" + j].Value = unidades.FECHA_FACTURA;
                        worksheet.Cells["S" + j].Value = unidades.COSTO_DEPRECIACION;
                        worksheet.Cells["T" + j].Value = unidades.STATUS_UNIDAD;
                        worksheet.Cells["U" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["U" + j].Value = unidades.FECHA_ASIGNACION;
                        worksheet.Cells["V" + j].Value = unidades.OBSERVACIONES;
                        worksheet.Cells["W" + j].Value = unidades.ODT;
                        worksheet.Cells["X" + j].Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                        worksheet.Cells["X" + j].Value = unidades.FECHA_ATENCION_PROVEEDOR;
                        j++;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
                return Ok(name);
            }
            catch (Exception ex)
            {
                return BadRequest(ex);
            }
        }
        [HttpPost("Negocios")]
        public async Task<ActionResult> Negocios(NegociosRequest negociosRequest, CancellationToken cancellationToken)
        {
            try
            {
                string folder = "C:\\inetpub\\wwwroot\\" + appname + "\\NEGOCIOS\\ARCHIVOS";
                string name = "Negocios_" + RandomString(10) + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + ".xlsx";
                string downloadUrl = "C:\\inetpub\\wwwroot\\" + appname + "\\NEGOCIOS\\ARCHIVOS\\" + name;
                Nullable<Int64> inull = null;
                FileInfo file = new FileInfo(Path.Combine(folder, name));

                var columHeaders = new string[]
                {
                    "DESC_NEGOCIO",
                    "NO_AFILIACION",
                    "ESTADO",
                    "ZONA",
                    "DIRECCION",
                    "COLONIA",
                    "POBLACION",
                    "CP",
                    "TELEFONO",
                    "RFC",
                    "RAZON SOCIAL",
                    "STATUS"
                };

                if (file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(Path.Combine(folder, name));
                }

                List<SpGetNegocios> procedure = new List<SpGetNegocios>();

                procedure = await context_.Query<SpGetNegocios>().FromSql("EXEC SP_GET_NEGOCIOS_API @p0", negociosRequest.searchText).ToListAsync();


                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Hoja1");
                    using (var cells = worksheet.Cells[1, 1, 1, 12])
                    {
                        cells.Style.Font.Bold = true;
                        cells.Style.Font.Color.SetColor(Color.White);
                        cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 128, 255));
                    }

                    for (var i = 0; i < columHeaders.Count(); i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columHeaders[i];
                    }
                    var j = 2;
                    foreach (var negocio in procedure)
                    {
                        worksheet.Cells["A" + j].Value = negocio.DESC_NEGOCIO;
                        worksheet.Cells["B" + j].Style.Numberformat.Format = "0";
                        worksheet.Cells["B" + j].Value = negocio.NO_AFILIACION == "" ? inull : Convert.ToInt64(negocio.NO_AFILIACION);
                        worksheet.Cells["C" + j].Value = negocio.DESC_ZONA;
                        worksheet.Cells["D" + j].Value = negocio.DESC_REGION;
                        worksheet.Cells["E" + j].Value = negocio.DIRECCION;
                        worksheet.Cells["F" + j].Value = negocio.COLONIA;
                        worksheet.Cells["G" + j].Value = negocio.POBLACION;
                        worksheet.Cells["H" + j].Value = negocio.CP;
                        worksheet.Cells["I" + j].Value = negocio.TELEFONO;
                        worksheet.Cells["J" + j].Value = negocio.RFC;
                        worksheet.Cells["K" + j].Value = negocio.RAZON_SOCIAL;
                        worksheet.Cells["L" + j].Value = negocio.STATUS;
                        j++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    //worksheet.Cells.LoadFromCollection(negocios, true);
                    package.Save();
                }
                return Ok(name);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
            .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}