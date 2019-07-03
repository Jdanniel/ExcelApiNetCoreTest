using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelApiNetCoreTest.DBModels;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace ExcelApiNetCoreTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportExcelController : ControllerBase
    {
        protected ELAVONTESTContext context_ = new ELAVONTESTContext();
        private static Random random = new Random();
        private readonly IConfiguration configuration;
        private readonly string appname;

        public ImportExcelController(IConfiguration iConfig)
        {
            configuration = iConfig;
            appname = configuration.GetValue<string>("MySettings:appname");
        }

        [HttpPost("transacciones")]
        public async Task<ActionResult> OnPostImportTransactions(IFormFile formFile, [FromForm] int IdUsuario, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest("FormFile esta vacio");
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("No es compatible con la extension del archivo");
            }
            try
            {
                String mn = DateTime.Now.AddMonths(1).Month.ToString();
                String yy = DateTime.Now.Year.ToString();

                BdCargasMasivas carga = new BdCargasMasivas()
                {
                    DescAttach = "Carga de Transacciones",
                    IdArchivoAttach = 2,
                    FechaCarga = DateTime.Now,
                    IdUsuarioAlta = IdUsuario,
                    Status = "Pendiente"
                };

                context_.BdCargasMasivas.Add(carga);
                await context_.SaveChangesAsync();


                int idcarga = carga.IdCargaMasiva;

                var list = new List<BdTransaccionesPaso>();

                using (var stream = new MemoryStream())
                {
                    await formFile.CopyToAsync(stream, cancellationToken);

                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            BdTransaccionesPaso paso = new BdTransaccionesPaso()
                            {
                                IdCarga = idcarga,
                                Month = Convert.ToInt32(mn),
                                Year = Convert.ToInt32(yy),
                                NoAfiliacion = Convert.ToInt32(worksheet.Cells[row, 1].Value.ToString().Trim()),
                                ApprovedcCount = Convert.ToInt32(worksheet.Cells[row, 2].Value.ToString().Trim()),
                                Declinedcount = Convert.ToInt32(worksheet.Cells[row, 3].Value.ToString().Trim())
                            };
                            list.Add(paso);
                        }
                        await context_.BdTransaccionesPaso.AddRangeAsync(list);
                        context_.SaveChanges();
                    }
                }

                await context_.Database.ExecuteSqlCommandAsync("EXEC SP_PROCESAR_CARGA_TRANSACCIONES @ID_CARGA",
                    new SqlParameter("@ID_CARGA", idcarga));

                return Ok(idcarga.ToString());
            }
            catch (Exception ex)
            {
                return BadRequest("A ocurrido un error: " + ex.ToString());
            }
        }

        [HttpPost("Bloqueos")]
        public async Task<ActionResult> OnPostImportLocks(IFormFile formFile, [FromForm] int IdUsuario, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest("FormFile esta vacio");
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("No es compatible con la extension del archivo");
            }
            try
            {
                String mn = DateTime.Now.AddMonths(1).Month.ToString();
                String yy = DateTime.Now.Year.ToString();

                BdCargasMasivas carga = new BdCargasMasivas()
                {
                    DescAttach = "Carga de Bloqueos",
                    IdArchivoAttach = 2,
                    FechaCarga = DateTime.Now,
                    IdUsuarioAlta = IdUsuario,
                    Status = "Pendiente"
                };

                context_.BdCargasMasivas.Add(carga);
                await context_.SaveChangesAsync();

                int idcarga = carga.IdCargaMasiva;

                var list = new List<BdBloqueosPaso>();

                using (var stream = new MemoryStream())
                {
                    await formFile.CopyToAsync(stream, cancellationToken);

                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var bloqueo = 0;
                            var celda = worksheet.Cells[row, 6].Value == null ? "No" : worksheet.Cells[row, 6].Value.ToString().Trim();
                            if (celda == "Si" || celda == "Sí")
                            {
                                bloqueo = 1;
                            }

                            BdBloqueosPaso paso = new BdBloqueosPaso()
                            {
                                IdCarga = idcarga,
                                Month = Convert.ToInt32(mn),
                                Year = Convert.ToInt32(yy),
                                NoAfiliacion = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                Cp = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                Proveedor = worksheet.Cells[row, 3].Value.ToString().Trim(),
                                Territorio = worksheet.Cells[row, 4].Value.ToString().Trim(),
                                TotalRollos = Convert.ToInt32(worksheet.Cells[row, 5].Value.ToString().Trim()),
                                Bloqueo = bloqueo,
                                Mensaje = worksheet.Cells[row, 7].Value == null ? null : worksheet.Cells[row, 7].Value.ToString().Trim()
                            };

                            list.Add(paso);
                        }
                        await context_.BdBloqueosPaso.AddRangeAsync(list);
                        await context_.SaveChangesAsync();
                    }
                }

                await context_.Database.ExecuteSqlCommandAsync("EXEC SP_PROCESAR_CARGA_BLOQUEOS @ID_CARGA",
                    new SqlParameter("@ID_CARGA", idcarga));

                return Ok(idcarga.ToString());
            }
            catch (Exception ex)
            {
                return BadRequest("A ocurrido un error: " + ex.ToString());
            }
        }

        [HttpPost("Cierres")]
        public async Task<ActionResult> OnPostImportClosures(IFormFile formFile, [FromForm] int IdUsuario, CancellationToken cancellationToken)
        {

            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest("FormFile esta vacio");
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("No es compatible con la extension del archivo");
            }
            var namefile = "Cierres_" + RandomString(10) + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + ".xlsx";
            CArchivos archivos = new CArchivos()
            {
                FechaAlta = DateTime.Now,
                NombreArchivo = namefile,
                IsCerradoExito = 1,
                Status = "PENDIENTE"
            };

            await context_.CArchivos.AddAsync(archivos);
            await context_.SaveChangesAsync();

            int idarchivo = archivos.IdArchivo;

            var list = new List<BdIngresoArchivosExito>();
            var listLayout = new List<BdExitoLayoutLog>();
            var listint = new List<int>();

            try
            {
                using (var stream = new FileStream(Path.Combine("C://inetpub//wwwroot//" + appname + "//SOLICITUDES/CIERRE_MASIVO//EXITO_RECHAZO//ARCHIVOS", namefile), FileMode.Create))
                {
                    await formFile.CopyToAsync(stream, cancellationToken);

                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rownCount = worksheet.Dimension.Rows;

                        var correo = worksheet.Cells[1, 15].Value == null ? null : worksheet.Cells[1, 15].Value.ToString().Trim();

                        for (int row = 2; row <= rownCount; row++)
                        {
                            var odt = worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString().Trim();
                            var minutos = worksheet.Cells[row, 6].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[row, 6].Value.ToString().Trim());
                            var statusMov = worksheet.Cells[row, 14].Value == null ? "ERROR" : worksheet.Cells[row, 14].Value.ToString().Trim();
                            var fecCierre = worksheet.Cells[row, 4].Value == null ? default(DateTime?) : DateTime.Parse(worksheet.Cells[row, 4].Value.ToString().Trim());
                            var statusArchivo = "PENDIENTE";
                            var descError = "";

                            if (minutos > 60)
                            {
                                statusMov = "ERROR";
                                statusArchivo = "PROCESADO";
                                descError = "EL CAMPO MINUTOS TIENE UN FORMATO INCORRECTO";
                            }

                            int idar = 0;

                            if (odt != null)
                            {
                                idar = context_.BdAr.Where(s => s.NoAr == odt.ToString()).Select(a => a.IdAr).FirstOrDefault();
                            }
                            else
                            {
                                statusMov = "ERROR";
                                statusArchivo = "PROCESADO";
                                descError = "EL CAMPO DE ODT ESTA VACIO";
                            }

                            BdIngresoArchivosExito exito = new BdIngresoArchivosExito()
                            {
                                IdArchivo = idarchivo,
                                IdAr = idar,
                                StatusArchivo = statusArchivo,
                                Aplicacion = worksheet.Cells[row, 8].Value == null ? null : worksheet.Cells[row, 8].Value.ToString().Trim(),
                                Version = worksheet.Cells[row, 9].Value == null ? null : worksheet.Cells[row, 9].Value.ToString().Trim(),
                                Caja = worksheet.Cells[row, 10].Value == null ? null : worksheet.Cells[row, 10].Value.ToString().Trim(),
                                Atiende = worksheet.Cells[row, 3].Value == null ? null : worksheet.Cells[row, 3].Value.ToString().Trim(),
                                FecCierre = fecCierre,
                                HoraCierre = worksheet.Cells[row, 5].Value == null ? null : worksheet.Cells[row, 5].Value.ToString().Trim(),
                                MinutoCierre = minutos.ToString(),
                                StatusMov = statusMov,
                                IdUsuarioAltaArchivo = IdUsuario,
                                FechaAltaArchivo = DateTime.Now,
                                DescError = descError
                            };
                            list.Add(exito);
                        }
                        await context_.BdIngresoArchivosExito.AddRangeAsync(list);
                        await context_.SaveChangesAsync();
                        listint = list.Select(x => x.IdArchivoMasivo).ToList();

                        for (int row = 2; row <= rownCount; row++)
                        {
                            var odt = worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString().Trim();
                            int idar = 0;

                            if (odt != null)
                            {
                                idar = context_.BdAr.Where(s => s.NoAr == odt.ToString()).Select(a => a.IdAr).FirstOrDefault();
                            }
                            var idarchivoMasivo0 = context_.BdIngresoArchivosExito.Where(x => listint.Contains(x.IdArchivoMasivo) && x.IdAr == idar).Select(a => a.IdArchivoMasivo).SingleAsync();

                            BdExitoLayoutLog layout = new BdExitoLayoutLog()
                            {
                                IdArchivoMasivo = await idarchivoMasivo0,
                                Odtexterna = worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString().Trim(),
                                CierreServicio = worksheet.Cells[row, 2].Value == null ? null : worksheet.Cells[row, 2].Value.ToString().Trim(),
                                Atiende = worksheet.Cells[row, 3].Value == null ? null : worksheet.Cells[row, 3].Value.ToString().Trim(),
                                FechaCierre = worksheet.Cells[row, 4].Value == null ? null : worksheet.Cells[row, 4].Value.ToString().Trim(),
                                HorasCierre = worksheet.Cells[row, 5].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[row, 5].Value.ToString().Trim()),
                                MinutoCierre = worksheet.Cells[row, 6].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[row, 6].Value.ToString().Trim()),
                                OtorganteVobo = worksheet.Cells[row, 7].Value == null ? null : worksheet.Cells[row, 7].Value.ToString().Trim(),
                                Aplicacion = worksheet.Cells[row, 8].Value == null ? null : worksheet.Cells[row, 8].Value.ToString().Trim(),
                                Version = worksheet.Cells[row, 9].Value == null ? null : worksheet.Cells[row, 9].Value.ToString().Trim(),
                                Caja = worksheet.Cells[row, 10].Value == null ? null : worksheet.Cells[row, 10].Value.ToString().Trim(),
                                OtorganteVoboRechazo = worksheet.Cells[row, 11].Value == null ? null : worksheet.Cells[row, 11].Value.ToString().Trim(),
                                Subrechazo = worksheet.Cells[row, 12].Value == null ? null : worksheet.Cells[row, 12].Value.ToString().Trim(),
                                IdCausaCancelacion = worksheet.Cells[row, 13].Value == null ? null : worksheet.Cells[row, 13].Value.ToString().Trim(),
                                Estatus = worksheet.Cells[row, 14].Value == null ? null : worksheet.Cells[row, 14].Value.ToString().Trim(),
                                Correo = correo,
                                TipoAtencion = worksheet.Cells[row, 16].Value == null ? null : worksheet.Cells[row, 16].Value.ToString().Trim(),
                                AmexSiNo = worksheet.Cells[row, 17].Value == null ? null : worksheet.Cells[row, 17].Value.ToString().Trim(),
                                ConclusionesAmex = worksheet.Cells[row, 18].Value == null ? null : worksheet.Cells[row, 18].Value.ToString().Trim(),
                                Idamx = worksheet.Cells[row, 19].Value == null ? null : worksheet.Cells[row, 19].Value.ToString().Trim(),
                                Afilamx = worksheet.Cells[row, 20].Value == null ? null : worksheet.Cells[row, 20].Value.ToString().Trim(),
                                Idrechazo = worksheet.Cells[row, 21].Value == null ? null : worksheet.Cells[row, 21].Value.ToString().Trim(),
                                TelefonoCampo = worksheet.Cells[row, 22].Value == null ? null : worksheet.Cells[row, 22].Value.ToString().Trim(),
                                ActReferencias = worksheet.Cells[row, 23].Value == null ? null : worksheet.Cells[row, 23].Value.ToString().Trim(),
                                Idcriteriocambio = worksheet.Cells[row, 24].Value == null ? null : worksheet.Cells[row, 24].Value.ToString().Trim(),
                                Discover = worksheet.Cells[row, 25].Value == null ? null : worksheet.Cells[row, 25].Value.ToString().Trim(),
                                Rollosinst = worksheet.Cells[row, 26].Value == null ? null : worksheet.Cells[row, 26].Value.ToString().Trim(),
                                IdArchivo = idarchivo,
                                IdAr = idar
                            };
                            listLayout.Add(layout);
                        }
                        await context_.BdExitoLayoutLog.AddRangeAsync(listLayout);
                        await context_.SaveChangesAsync();
                    }
                }
                return Ok(idarchivo);
            }
            catch (Exception ex)
            {
                return BadRequest("A ocurrido un error: " + ex.ToString());
            }
        }

        [HttpPost("Transferencias")]
        public async Task<ActionResult> OnPostImportTranfers(IFormFile formFile, [FromForm] int ID_TIPO_RESPONSABLE_O, [FromForm] int ID_TIPO_RESPONSABLE_D, [FromForm] int ID_RESPONSABLE_O, [FromForm] int ID_RESPONSABLE_D, [FromForm] int ID_TRANSFERENCIA, [FromForm] int ID_USUARIO, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest("FormFile esta vacio");
            }
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("No es compatible con la extension del archivo");
            }

            var namefile = "Tranferencias_" + RandomString(10) + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + ".xlsx";

            try
            {
                var transferencia = await context_.BdTransferencias.FirstOrDefaultAsync(x => x.IdTransferencia == ID_TRANSFERENCIA);
                var list = new List<BdTransferenciasPaso>();
                if (transferencia != null)
                {
                    transferencia.SystemFilename = namefile;
                    transferencia.UserFilename = formFile.FileName;

                    context_.BdTransferencias.Update(transferencia);
                    await context_.SaveChangesAsync();

                    using (var stream = new FileStream(Path.Combine("C://inetpub//wwwroot//" + appname + "//ALMACEN/TRANSFERENCIAS//ARCHIVOS", namefile), FileMode.Create))
                    {
                        await formFile.CopyToAsync(stream, cancellationToken);

                        using (var package = new ExcelPackage(stream))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                            var rownCount = worksheet.Dimension.Rows;
                            for (int row = 2; row <= rownCount; row++)
                            {
                                var noserie = worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString().Trim();
                                if (noserie != "")
                                {
                                    BdTransferenciasPaso transferenciasPaso = new BdTransferenciasPaso()
                                    {
                                        IdTransferencia = ID_TRANSFERENCIA,
                                        NoSerie = worksheet.Cells[row, 1].Value == null ? null : worksheet.Cells[row, 1].Value.ToString().Trim(),
                                        IdTipoResponsableOrigen = ID_TIPO_RESPONSABLE_O,
                                        IdResponsableOrigen = ID_RESPONSABLE_O,
                                        IdTipoResponsableDestino = ID_TIPO_RESPONSABLE_D,
                                        IdResponsableDestino = ID_RESPONSABLE_D,
                                        IdUsuario = ID_USUARIO,
                                        FecAlta = DateTime.Now
                                    };
                                    list.Add(transferenciasPaso);
                                }
                            }
                            if (list.Count > 0)
                            {
                                await context_.BdTransferenciasPaso.AddRangeAsync(list);
                                await context_.SaveChangesAsync();
                            }
                        }
                    }
                }
                else
                {
                    return BadRequest("La transferencia no existe");
                }
                return Ok(ID_TRANSFERENCIA);
            }
            catch (Exception ex)
            {
                return BadRequest("A ocurrido un error: " + ex.ToString());
            }

        }

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
            .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }
    }
}