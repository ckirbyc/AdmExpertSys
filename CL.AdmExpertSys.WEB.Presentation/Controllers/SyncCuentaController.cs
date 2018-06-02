using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Services.Protocols;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class SyncCuentaController : BaseController
    {
        // GET: SyncCuenta
        private static Thread _hiloEjecucion;

        protected EstadoCuentaUsuarioFactory EstadoCuentaUsuarioFactory;
        protected HomeSysWebFactory HomeSysWebFactory;

        public SyncCuentaController()
        {
        }

        public SyncCuentaController(EstadoCuentaUsuarioFactory estadoCuentaUsuarioFactory)
        {
            EstadoCuentaUsuarioFactory = estadoCuentaUsuarioFactory;
        }
        
        public ActionResult Index()
        {
            try
            {
                ViewBag.EstadoSync = HiloEstadoSincronizacion.EsSincronizacion();
                //Obtener todos los usuarios No Sincronizados
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoSync();
                return View(listaEstUsr);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Sincronizacion. Si el problema persiste contacte a soporte IT" });
            }            
        }

        [HttpPost]
        public ActionResult SincronizarCuentas()
        {
            try
            {
                var listaEstUsr = EstadoCuentaUsuarioFactory.GetEstadoCuentaUsuarioNoSync();

                var listaEstCuentaVmHilo = new List<object>
                {
                    listaEstUsr
                };

                _hiloEjecucion = new Thread(InciarProcesoHiloSincronizarCuenta);
                _hiloEjecucion.Start(listaEstCuentaVmHilo);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = string.Empty
                    }
                };
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = false,
                        Error = ex.Message
                    }
                };
            }
        }

        [SoapDocumentMethod(OneWay = true)]
        public void InciarProcesoHiloSincronizarCuenta(object estadoCuentaHilo)
        {
            HiloEstadoSincronizacion.ActualizarEstadoSync(true);

            HomeSysWebFactory = new HomeSysWebFactory();
            try
            {
                HomeSysWebFactory.ForzarDirSync();
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
            }           

            var estadoCuentaLista = (List<EstadoCuentaUsuarioVm>)estadoCuentaHilo.CastTo<List<object>>()[0];

            Task.Delay(TimeSpan.FromSeconds(120)).Wait();

            foreach (var estUsr in estadoCuentaLista)
            {
                if (HomeSysWebFactory.ExisteUsuarioPortal(estUsr.Correo.Trim()))
                {
                    estUsr.Sincronizado = true;
                    try
                    {
                        HiloEstadoCuentaUsuario.ActualizarEstadoCuentaUsuario(estUsr);
                    }
                    catch(Exception ex)
                    {
                        Utils.LogErrores(ex);
                    }                    
                }
            }

            HiloEstadoSincronizacion.ActualizarEstadoSync(false);
        }
    }
}