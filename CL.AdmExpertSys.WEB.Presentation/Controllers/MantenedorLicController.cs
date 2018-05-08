using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class MantenedorLicController : BaseController
    {
        // GET: MantenedorLic
        protected MantenedorLicenciaFactory MantenedorLicenciaFactory;
        protected Common CommonFactory;

        public MantenedorLicController()
        {
        }

        public MantenedorLicController(MantenedorLicenciaFactory mantenedorLicenciaFactory)
        {
            MantenedorLicenciaFactory = mantenedorLicenciaFactory;
        }
        public ActionResult Index()
        {
            try
            {
                //Obtener todos los datos de la tabla mantenedores de licencias
                var listaMantLic = MantenedorLicenciaFactory.GetAllMantenedorLicencia();
                return View(listaMantLic);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Mantenedor Licencia. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult Create()
        {
            return View(MantenedorLicenciaFactory.GetCreateView());
        }

        [HttpPost]
        public ActionResult Create(MantenedorLicenciaVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                if (!ModelState.IsValid)
                {
                    model = MantenedorLicenciaFactory.GetCreateView();
                    return View(model);
                }

                MantenedorLicenciaFactory.CrearMantenedorLicencia(model);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar guardar registro de Mantenedor Licencia. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult Edit(decimal id)
        {
            return View(MantenedorLicenciaFactory.GetMantenedorLicencia(id));
        }

        [HttpPost]
        public ActionResult Edit(MantenedorLicenciaVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                if (!ModelState.IsValid)
                {
                    model = MantenedorLicenciaFactory.GetCreateView();
                    return View(model);
                }
                model.Vigente = true;
                MantenedorLicenciaFactory.ActualizaMantenedorLicencia(model);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar guardar registro de Mantenedor Licencia. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult Details(int id)
        {
            return View(MantenedorLicenciaFactory.GetMantenedorLicencia(id));
        }

        public ActionResult Delete(int id)
        {
            return View(MantenedorLicenciaFactory.GetMantenedorLicencia(id));
        }

        [HttpPost]
        public ActionResult Delete(MantenedorLicenciaVm model)
        {
            try
            {
                var objVm = MantenedorLicenciaFactory.GetMantenedorLicencia(model.IdMantenedorLicencia);
                objVm.Vigente = false;
                MantenedorLicenciaFactory.ActualizaMantenedorLicencia(objVm);

                return RedirectToAction("Index");
            }
            catch (DbUpdateException ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Los datos que intenta eliminar esta siendo utilizado en algún registro dentro de AdmExpertSys. Por favor reinténtelo más tarde." });
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar eliminar registro de Mantenedor Licencia. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult DescargarPlanilla()
        {
            try
            {                
                return View();
            }
            
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al cargar página para Descargar Planilla. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult CargarPlanilla(string error)
        {
            try
            {
                ViewData["errorExcel"] = error;
                var vm = new MantenedorLicenciaFileVm();
                return View(vm);
            }

            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al cargar página para Cargar Planilla. Por favor reinténtelo más tarde." });
            }
        }

        public FileStreamResult ExportarExcel()
        {
            try
            {               
                var libro = MantenedorLicenciaFactory.ExportarArchivoExcel();
                var memoryStream = new MemoryStream();
                libro.SaveAs(memoryStream);
                libro.Dispose();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                memoryStream.Flush();
                memoryStream.Position = 0;

                var nombreArchivo = string.Format("PlanillaFormatoLicencia.xlsx");
                return File(memoryStream, "MantendorLic", nombreArchivo);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return null;
            }
        }

        [HttpPost]
        public ActionResult ImportarExcel(MantenedorLicenciaFileVm model)
        {            
            var archivo = model.Archivo;

            if (model == null) throw new ArgumentNullException("model");

            if (archivo != null && archivo.ContentLength > 0)
            {
                // Validaciones del Archivo
                var extension = Path.GetExtension(archivo.FileName);
                var err = false;
                string msj;

                if (extension != null)
                {
                    var ext = extension.Replace(".", string.Empty);
                    var extensionesValidas = @"xlsx";
                    var extErr = extensionesValidas.Contains(ext);
                    msj = @"El archivo posee una extensión no válida, seleccione un archivo de tipo Excel (.xlsx).";

                    if (!extErr) err = true;
                }
                else
                {
                    err = true;
                    msj = @"No se pudo validar la extensión del archivo, por favor seleccione un archivo Excel válido (.xlsx).";
                }
                if (err)
                {
                    return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = msj });
                }

                CommonFactory = new Common();
                var rutaTemp = CommonFactory.GetAppSetting("RutaExcel");
                var exiRut = Directory.Exists(rutaTemp);
                if (!exiRut)
                {
                    try
                    {
                        Directory.CreateDirectory(rutaTemp);
                    }
                    catch (Exception ex)
                    {
                        Utils.LogErrores(ex);
                        return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = "Error al intentar Procesar archivo Excel, la ruta para generación del Archivo no existe y no es posible crearla. Por favor reinténtelo más tarde." });
                    }
                }

                var nombreArchivoNue = string.Format("FormatoLicencia_{0}{1}", DateTime.Now.ToString("HHmmss"), extension);
                rutaTemp = rutaTemp + nombreArchivoNue;
                archivo.SaveAs(rutaTemp);

                try
                {
                    //Inciar Proceso Carga Archivo Excel
                    var listaMantLicVm = new List<MantenedorLicenciaVm>();
                    try
                    {
                        listaMantLicVm = MantenedorLicenciaFactory.ImportarExcelToLista(rutaTemp);
                    }
                    catch (FormatException ex)
                    {
                        return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = ex.Message });
                    }

                    if (listaMantLicVm.Any())
                    {
                        var resultOp = MantenedorLicenciaFactory.TruncateTablaMantenedorLicencia();
                        if (resultOp)
                        {
                            resultOp = MantenedorLicenciaFactory.InsertarDatoMasivoMantenedorLicencia(listaMantLicVm);
                        }

                        if (resultOp)
                        {                            
                            return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = "Proceso carga planilla Excel de Licencias, terminado exitosamente." });
                        }
                        return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = "Error en el proceso de guardar valores Licencias. Reintente más tarde." });
                    }
                    return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = "El documento se encuentra vacío." });
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                    return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = ex.Message });
                }
            }
            return RedirectToAction("CargarPlanilla", "MantenedorLic", new { error = @"Debe seleccionar un archivo." });
        }
    }
}