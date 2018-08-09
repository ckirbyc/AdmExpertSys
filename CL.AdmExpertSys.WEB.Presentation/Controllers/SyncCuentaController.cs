using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Application.CommonLib;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.Models;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Threading;
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
                var usuarioModificacion = SessionViewModel.Usuario.Nombre.Trim();
                var listaEstCuentaVmHilo = new List<object>
                {
                    listaEstUsr,
                    usuarioModificacion
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
            var usuarioModificacion = (string)estadoCuentaHilo.CastTo<List<object>>()[1];
            try
            {               
                HiloEstadoSincronizacion.ActualizarEstadoSync(true, usuarioModificacion, "S");

                HomeSysWebFactory = new HomeSysWebFactory();
                try
                {
                    HomeSysWebFactory.ForzarDirSync();
                }
                catch (Exception ex)
                {
                    HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "S");
                    Utils.LogErrores(ex);
                }

                var estadoCuentaLista = (List<EstadoCuentaUsuarioVm>)estadoCuentaHilo.CastTo<List<object>>()[0];

                //Task.Delay(TimeSpan.FromSeconds(120)).Wait();
                var comm = new Common();
                var intentoEstadoSync = Convert.ToInt64(comm.GetAppSetting("IntentoEstadoSync"));

                foreach (var estUsr in estadoCuentaLista)
                {
                    var upnPrefijo = estUsr.CuentaAd.Trim() + estUsr.Dominio.Trim();
                    for (int i = 1; i <= intentoEstadoSync; i++)
                    {
                        if (HomeSysWebFactory.ExisteUsuarioPortal(upnPrefijo))
                        {
                            estUsr.Sincronizado = true;
                            try
                            {
                                HiloEstadoCuentaUsuario.ActualizarEstadoCuentaUsuario(estUsr);
                                break;
                            }
                            catch (Exception ex)
                            {
                                HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "S");
                                Utils.LogErrores(ex);
                                break;
                            }
                        }
                    }                    
                }
                HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "S");
            }
            catch (Exception ex)
            {
                HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "S");
                var msgError = @"Error en proceso asincronico Syncronizar : " + ex.Message;
                var exNew = new Exception(msgError);
                Utils.LogErrores(exNew);                
            }            
        }
    }
}