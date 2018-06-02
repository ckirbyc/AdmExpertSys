using AutoMapper;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using ClosedXML.Excel;
using Pragma.Commons.Data.Patterns.Specification;
using Pragma.Commons.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class MantenedorLicenciaFactory
    {
        protected IMantenedorLicenciaService MantenedorLicenciaService;
        protected ITransversalService TransversalService;
        protected IRolCargoService RolCargoService;
        protected ILicenciaO365Service LicenciaO365Service;

        public MantenedorLicenciaFactory(
            IMantenedorLicenciaService mantenedorLicenciaService,
            ITransversalService transversalService,
            IRolCargoService rolCargoService,
            ILicenciaO365Service licenciaO365Service)
        {
            MantenedorLicenciaService = mantenedorLicenciaService;
            TransversalService = transversalService;
            RolCargoService = rolCargoService;
            LicenciaO365Service = licenciaO365Service;
        }

        public IList<MantenedorLicenciaVm> GetAllMantenedorLicencia()
        {
            var espec = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.Vigente);
            var objBd = MantenedorLicenciaService.AllMatching(espec);

            var vm = Mapper.Map<IQueryable<MANTENEDOR_LICENCIA>, IList<MantenedorLicenciaVm>>(objBd);
            return vm.ToList();
        }

        public MantenedorLicenciaVm GetCreateView()
        {
            var mantLicVm = new MantenedorLicenciaVm
            {
                RolCargoSelectListItem = TransversalService.GetSelectRolCarga(),
                LicenciaSelectListItem = TransversalService.GetSelectLicencia()
            };
            return mantLicVm;
        }

        public void CrearMantenedorLicencia(MantenedorLicenciaVm tipo)
        {
            var mantLic = new MANTENEDOR_LICENCIA
            {                
                Codigo = tipo.Codigo,
                RolCargoId = tipo.RolCargoId,
                LicenciaId = tipo.LicenciaId,
                Vigente = true
            };
            MantenedorLicenciaService.Create(mantLic);
        }

        public MantenedorLicenciaVm GetMantenedorLicencia(decimal id)
        {
            var espec = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.IdMantenedorLicencia == id);
            var objBd = MantenedorLicenciaService.AllMatching(espec);

            var objVm = Mapper.Map(objBd.FirstOrDefault(), new MantenedorLicenciaVm());
            objVm.RolCargoSelectListItem = TransversalService.GetSelectRolCarga();
            objVm.LicenciaSelectListItem = TransversalService.GetSelectLicencia();

            return objVm;
        }

        public void ActualizaMantenedorLicencia(MantenedorLicenciaVm tipo)
        {
            var espec = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.IdMantenedorLicencia == tipo.IdMantenedorLicencia);
            var objBd = MantenedorLicenciaService.AllMatching(espec).FirstOrDefault();

            if (objBd != null)
            {
                objBd.Codigo = tipo.Codigo;
                objBd.LicenciaId = tipo.LicenciaId;
                objBd.RolCargoId = tipo.RolCargoId;                
                objBd.Vigente = tipo.Vigente;

                MantenedorLicenciaService.Update(objBd);
            }               
        }

        public bool ExisteCodigoLicencia(string codigo)
        {
            try
            {                
                var espec = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.Codigo == codigo.Trim());
                var mantLicBd = MantenedorLicenciaService.AllMatching(espec);

                if (mantLicBd.Any()) {
                    return true;
                }

                return false;                
            }
            catch (Exception ex) {
                Utils.LogErrores(ex);
                return false;
            }
        }

        public MantenedorLicenciaVm ObtenerLicenciaCodigo(string codigo)
        {
            var mantLic = new MantenedorLicenciaVm();
            try
            {                
                var espec = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.Codigo == codigo.Trim());
                var mantLicBd = MantenedorLicenciaService.AllMatching(espec);

                if (mantLicBd.Any())
                {
                    var retorno = Mapper.Map(mantLicBd.FirstOrDefault(), mantLic);
                    return retorno;
                }

                return mantLic;
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return mantLic;
            }
        }

        public XLWorkbook ExportarArchivoExcel()
        {
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                using (var hojaLic = workbook.Worksheets.Add("Datos Licencia"))
                {
                    hojaLic.Cell(1, 1).Value = "Código";
                    hojaLic.Cell(1, 1).Style.Font.FontColor = XLColor.White;
                    hojaLic.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.Red;
                    hojaLic.Cell(1, 1).Style.Font.Bold = true;

                    hojaLic.Cell(1, 2).Value = "Rol Cargo (Seleccione)";
                    hojaLic.Cell(1, 2).Style.Font.FontColor = XLColor.White;
                    hojaLic.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.Red;
                    hojaLic.Cell(1, 2).Style.Font.Bold = true;

                    hojaLic.Cell(1, 3).Value = "Licencia (Seleccione)";
                    hojaLic.Cell(1, 3).Style.Font.FontColor = XLColor.White;
                    hojaLic.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.Red;
                    hojaLic.Cell(1, 3).Style.Font.Bold = true;

                    var especMantLic = new DirectSpecification<MANTENEDOR_LICENCIA>(x => x.Vigente);
                    var mantLicBd = MantenedorLicenciaService.AllMatching(especMantLic);
                    var listaMantLicOrdenada = mantLicBd.OrderBy(x => x.IdMantenedorLicencia);
                    var listaExcelMantLic = new List<string[]>(listaMantLicOrdenada.Count());
                    var fila = 0;

                    //Hoja Parametros
                    using (var hojaParam = workbook.Worksheets.Add("Parametros").Hide())
                    {
                        //var selecRolCargo = "[0]--Seleccione--";
                        var especRol = new DirectSpecification<ROL_CARGO>(x => x.Vigente);
                        var rolCargoBd = RolCargoService.AllMatching(especRol);
                        var rolCargoVm = Mapper.Map<IQueryable<ROL_CARGO>, IList<RolCargoVm>>(rolCargoBd);

                        var rolCargoAll = (from a in rolCargoVm.OrderBy(x => x.Nombre)
                                           select "[" + a.IdRolCargo.ToString() + "]-" + a.Nombre).ToList();
                        var totalRolCargoAll = rolCargoAll.Count;
                        hojaParam.Cell(1, 1).InsertData(rolCargoAll);

                        var especLic = new DirectSpecification<LICENCIAS_O365>(x => x.Vigente);
                        var licenciaBd = LicenciaO365Service.AllMatching(especLic);
                        var licenciaVm = Mapper.Map<IQueryable<LICENCIAS_O365>, IList<LicenciaO365Vm>>(licenciaBd);
                        //var selecLicencia = "[0]--Seleccione--";
                        var licenciaAll = (from a in licenciaVm.OrderBy(x => x.Nombre)
                                           select "[" + a.IdLicencia.ToString() + "]-" + a.Nombre).ToList();
                        var totalLicenciaAll = licenciaAll.Count;
                        hojaParam.Cell(1, 2).InsertData(licenciaAll);

                        var rangoListaRolCargo = hojaParam.Range("A1:A" + totalRolCargoAll);
                        var rangoListaLic = hojaParam.Range("B1:B" + totalLicenciaAll);

                        hojaLic.Range("B2:B" + listaMantLicOrdenada.Count() + 1).Style.Font.Bold = true;
                        hojaLic.Range("B2:B" + +listaMantLicOrdenada.Count() + 1).SetDataValidation().List(rangoListaRolCargo);

                        hojaLic.Range("C2:C" + listaMantLicOrdenada.Count() + 1).Style.Font.Bold = true;
                        hojaLic.Range("C2:C" + +listaMantLicOrdenada.Count() + 1).SetDataValidation().List(rangoListaLic);

                        hojaParam.Columns().AdjustToContents();
                    }

                    foreach (var mantLic in listaMantLicOrdenada) {
                        var listaMantLic = new string[3];
                        listaMantLic[0] = mantLic.Codigo.ToString();
                        listaMantLic[1] = $"[{mantLic.RolCargoId.ToString()}]-{mantLic.ROL_CARGO.Nombre}";
                        listaMantLic[2] = $"[{mantLic.LicenciaId.ToString()}]-{mantLic.LICENCIAS_O365.Nombre}";
                        listaExcelMantLic.Add(listaMantLic);
                        fila++;
                    }

                    hojaLic.Cell(2, 1).InsertData(listaExcelMantLic.AsEnumerable());
                    hojaLic.Columns().AdjustToContents();
                }               

                return workbook;
            }            
        }

        public List<MantenedorLicenciaVm> ImportarExcelToLista(string rutaArchivo)
        {
            try
            {
                var listaVm = new List<MantenedorLicenciaVm>();

                //Procesar archivo
                using (var workbook = new XLWorkbook(rutaArchivo))
                {
                    //hoja 1 es donde se encuentran los datos
                    using (var hojaMantLic = workbook.Worksheet(1))
                    {
                        int numFila = 1;

                        var tabla = hojaMantLic.Range(hojaMantLic.Row(hojaMantLic.FirstRowUsed().RowUsed().RowNumber()).FirstCell().Address,
                            hojaMantLic.LastCellUsed().Address).RangeUsed().AsTable();

                        try
                        {
                            foreach (var regFila in tabla.DataRange.Rows())
                            {
                                var codigoLic = regFila.Field(0).GetString();
                                var rolCargoIdString = regFila.Field(1).GetString().Split('-')[0].Replace("[", "").Replace("]", "");
                                var licenciaIdString = regFila.Field(2).GetString().Split('-')[0].Replace("[", "").Replace("]", "");

                                if (!string.IsNullOrEmpty(rolCargoIdString) &&
                                    !string.IsNullOrEmpty(licenciaIdString))
                                {
                                    //Procesar 
                                    var objVm = new MantenedorLicenciaVm
                                    {
                                        Codigo = codigoLic.Trim(),
                                        RolCargoId = Convert.ToDecimal(rolCargoIdString),
                                        LicenciaId = Convert.ToDecimal(licenciaIdString),
                                        Vigente = true
                                    };
                                    listaVm.Add(objVm);
                                }
                                else {
                                    var msgError = string.Empty;
                                    if (string.IsNullOrEmpty(rolCargoIdString))
                                        msgError =
                                            $"Revisar en fila {numFila}, debe existir el código del Rol Cargo [número código]";
                                    if (string.IsNullOrEmpty(licenciaIdString))
                                        msgError = msgError + " " +
                                            $"Revisar en fila {numFila}, debe existir el código de la Licencia [número código]";
                                    if (!string.IsNullOrEmpty(msgError))
                                    {                                                                            
                                        throw new FormatException(msgError);
                                    }
                                }

                                numFila++;
                                regFila.Clear();
                            }
                        }
                        catch (FormatException ex)
                        {
                            Utils.LogErrores(ex);
                            if (ex.Message.Contains("Revisar en fila"))
                                throw new FormatException(ex.Message);

                            throw new FormatException(
                                $"Error al leer el archivo, asegúrese de que los campos están rellenados correctamente en la fila {numFila} del documento. Fijarse en los ejemplos de cada columna.");
                        }
                    }                        
                }                                

                return listaVm;
            }
            catch (FormatException ex)
            {
                throw new FormatException(ex.Message);
            }
        }

        public bool TruncateTablaMantenedorLicencia()
        {
            try
            {
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var objCtx = ((System.Data.Entity.Infrastructure.IObjectContextAdapter)entityContext).ObjectContext)
                    {
                        objCtx.ExecuteStoreCommand("TRUNCATE TABLE [MANTENEDOR_LICENCIA]");
                        return true;
                    }                        
                }                    
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new Exception(ex.Message);
            }
        }

        public bool InsertarDatoMasivoMantenedorLicencia(List<MantenedorLicenciaVm> listaVm)
        {
            try
            {
                var listaBd = Mapper.Map<List<MantenedorLicenciaVm>, IList<MANTENEDOR_LICENCIA>>(listaVm);
                using (var entityContext = new AdmSysWebEntities())
                {
                    using (var dbContextTransaction = entityContext.Database.BeginTransaction())
                    {
                        entityContext.MANTENEDOR_LICENCIA.AddRange(listaBd);
                        entityContext.SaveChanges();
                        dbContextTransaction.Commit();
                        return true;
                    }                        
                }                    
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                throw new Exception(ex.Message);
            }
        }
    }
}
