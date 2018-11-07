using CL.AdmExpertSys.Web.Infrastructure.Helpers;
using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Core.Domain.Enums;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Thread;
using CL.AdmExpertSys.WEB.Presentation.Models;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using System.Web.Mvc;
using System.Web.Services.Protocols;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class HomeController : BaseController
    {
        private static Thread _hiloEjecucion;

        protected HomeSysWebFactory HomeSysWebFactory;
        protected LogInfoFactory LogInfoFactory;
        protected EstadoCuentaUsuarioFactory EstadoCuentaUsuarioFactory;
        protected MantenedorLicenciaFactory MantenedorLicenciaFactory;
        protected UsuarioFactory UsuarioFactory;

        public HomeController()
        {
        }

        public HomeController(LogInfoFactory logInfoFactory, 
            EstadoCuentaUsuarioFactory estadoCuentaUsuarioFactory,
            MantenedorLicenciaFactory mantenedorLicenciaFactory,
            UsuarioFactory usuarioFactory)
        {
            LogInfoFactory = logInfoFactory;
            EstadoCuentaUsuarioFactory = estadoCuentaUsuarioFactory;
            MantenedorLicenciaFactory = mantenedorLicenciaFactory;
            UsuarioFactory = usuarioFactory;
        }

        [AllowNotAutenticate]
        public ActionResult Index(LoginVm model, string mensajeError)
        {            
            try
            {
                ViewBag.MensajeError = mensajeError;
                model.UtilizarAutenticacion = Convert.ToBoolean(ConfigurationManager.AppSettings["UtilizarAutenticacion"]);
                return View(model);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página principal. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult HomeSysWeb()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();
                //ViewBag.EstructuraArbolAd = null;
                if (SessionViewModel.EstructuraArbolAd != null)
                {
                    ViewBag.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }
                else
                {
                    var listaOuAd = HomeSysWebFactory.GetListaOu();
                    var estructura = HomeSysWebFactory.GetArquitecturaArbolAd(listaOuAd);
                    SessionViewModel.EstructuraArbolAd = estructura;
                    ViewBag.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }                

                return View(HomeSysWebFactory.ObtenerVistaHomeSysWeb());
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página principal. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult EditCuenta()
        {
            try
            {
                HomeSysWebFactory = new HomeSysWebFactory();                
                if (SessionViewModel.EstructuraArbolAd != null)
                {
                    ViewBag.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }
                else
                {
                    var listaOuAd = HomeSysWebFactory.GetListaOu();
                    var estructura = HomeSysWebFactory.GetArquitecturaArbolAd(listaOuAd);
                    SessionViewModel.EstructuraArbolAd = estructura;
                    ViewBag.EstructuraArbolAd = SessionViewModel.EstructuraArbolAd;
                }

                return View(HomeSysWebFactory.ObtenerVistaHomeSysWeb());
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Editar Cuenta Usuario. Si el problema persiste contacte a soporte IT" });
            }
        }

        /// <summary>
        /// Controler valida credenciales del usuario al ingresar al sistema.
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [AllowNotAutenticate]
        [HttpPost]
        public ActionResult ValidaLogin(LoginVm model)
        {
            try
            {
                System.Web.HttpContext.Current.Session["AdmExpertSys"] = "AdmExpertSys";
                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Logeo usuario al sistema",
                    UserInfo = model.NombreUsuario,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.Login.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }                                
                SessionViewModel.Usuario = new UsuarioVm
                {
                    Id = 1,
                    Nombre = model.NombreUsuario,
                    Email = string.Empty,
                    Rut = 0,
                    Dv = string.Empty
                };

                if (!model.UtilizarAutenticacion)
                {
                    model.EstaAutenticado = true;
                    return RedirectToAction("HomeSysWeb", "Home");
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var usuario = model.NombreUsuario.ToLower().Trim();
                var clave = model.ClaveUsuario.Trim();

                if (HomeSysWebFactory.ValidarCredencialesUsuarioAd(usuario, clave))
                {
                    //Obtener Perfil Usuario
                    var objUsrPerfil = UsuarioFactory.GetUsuarioPerfilByNombreCta(usuario);
                    if (objUsrPerfil != null)
                    {
                        SessionViewModel.Usuario.PerfilId = objUsrPerfil.PerfilId;
                        SessionViewModel.Usuario.EsAdm = objUsrPerfil.EsAdm;
                    }
                    else
                    {
                        return RedirectToAction("IndexLogin", "Error", new { mensajeError = "Usuario no tiene asignado un Perfil. Favor contacte a soporte IT" });
                    }
                    model.EstaAutenticado = true;
                    return RedirectToAction("HomeSysWeb", "Home");
                }

                model.EstaAutenticado = false;
                return RedirectToAction("IndexLogin", "Error", new { mensajeError = "Error, usuario o contraseña no corresponden. Si el problema persiste contacte a soporte IT" });
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("IndexLogin", "Error", new { mensajeError = "Error al iniciar sesión. Si el problema persiste contacte a soporte IT" });
            }
        }

        [HttpPost]
        public ActionResult VerificarUsuario(string nombreUsuario)
        {
            var varSession = true;
            try
            {
                if (string.IsNullOrEmpty(nombreUsuario)) throw new ArgumentNullException("nombreUsuario");

                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {                    
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            DatosUsuario = string.Empty,
                            CodigoLicencia = string.Empty,
                            Clave = string.Empty,
                            Session = varSession,
                            EstadoSync = false
                        }
                    };
                }

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Verifica existencia usuario en el AD",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.VerificarExistenciaUsuario.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }
                
                HomeSysWebFactory = new HomeSysWebFactory();                
                var usuarioAd = HomeSysWebFactory.ObtenerUsuarioExistente(nombreUsuario.Trim());
                bool chequear = usuarioAd != null;
                var codigoLic = string.Empty;
                var claveCta = string.Empty;
                var estadoSync = false;

                if (chequear) {
                    codigoLic = EstadoCuentaUsuarioFactory.GetCodigoLicenciaByUsuario(nombreUsuario.Trim());
                    claveCta = EstadoCuentaUsuarioFactory.GetClaveCuentaByUsuario(nombreUsuario.Trim());
                    estadoSync = HiloEstadoSincronizacion.EsSincronizacion();
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = chequear,
                        DatosUsuario = usuarioAd,
                        CodigoLicencia = codigoLic.ToString(),
                        Clave = claveCta,
                        Session = varSession,
                        EstadoSync = estadoSync
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
                        DatosUsuario = string.Empty,
                        CodigoLicencia = string.Empty,
                        Clave = string.Empty,
                        Session = varSession,
                        EstadoSync = false
                    }
                };
            }
        }
        /// <summary>
        /// Guardar Usuario en AD
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GuardarSincronizarUsuario(HomeSysWebVm model)
        {
            var varSession = true;
            var codRuta = string.Empty;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {                    
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession,
                            CodigoRuta = codRuta
                        }
                    };
                }

                if (string.IsNullOrEmpty(model.Nombres)) throw new ArgumentException("Ingresar nombres del usuario");
                if (string.IsNullOrEmpty(model.Apellidos)) throw new ArgumentException("Ingresar apellidos del usuario");
                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");
                if (string.IsNullOrEmpty(model.Correo)) throw new ArgumentException("Asignar correo");
                if (string.IsNullOrEmpty(model.UpnPrefijo)) throw new ArgumentException("Ingresar Dominio");
                if (string.IsNullOrEmpty(model.Descripcion)) throw new ArgumentException("Ingresar Descripción");
                if (string.IsNullOrEmpty(model.CodigoLicencia)) throw new ArgumentException("Ingresar Código Licencia");
                if (model.ExisteUsuario == false)
                {
                    if (string.IsNullOrEmpty(model.Clave)) throw new ArgumentException("Ingresar clave del usuario");
                }
                else
                {
                    model.Clave = @"$aaa123";
                }

                if (string.IsNullOrEmpty(model.PatchOu) && string.IsNullOrEmpty(model.CodigoRutaAd))
                {
                    throw new ArgumentException("Seleccionar unidad organizativa o código ruta AD.");
                }
                else if (!string.IsNullOrEmpty(model.PatchOu) && !string.IsNullOrEmpty(model.CodigoRutaAd))
                {
                    throw new ArgumentException("Solo seleccionar unidad organizativa o código ruta AD. No pueden ser ambas.");
                }
                else if (string.IsNullOrEmpty(model.PatchOu) && !string.IsNullOrEmpty(model.CodigoRutaAd))
                {
                    //Validar que el código de la ruta ingresa exista en el AD.
                    HomeSysWebFactory = new HomeSysWebFactory();
                    codRuta = HomeSysWebFactory.ValidarCodigoRutaOuAd(model.CodigoRutaAd.Trim());
                    if (string.IsNullOrEmpty(codRuta))
                    {
                        throw new ArgumentException("El código ruta OU ingresado no existe en el AD.");
                    }
                    else
                    {
                        model.PatchOu = codRuta;
                    }
                }

                //if (string.IsNullOrEmpty(model.PatchOu))
                //{
                //    if (model.ExisteUsuario == false)
                //    {
                //        throw new ArgumentException("Seleccionar unidad organizativa");
                //    }
                //}
                
                //Validar que el código ingresado exista
                if (!MantenedorLicenciaFactory.ExisteCodigoLicencia(model.CodigoLicencia.Trim())) {
                    throw new ArgumentException("El código licencia ingresado no existe en la base de datos.");
                }

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Guarda Usuario en el AD",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.Guardar.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }

                var logSync = new LogInfoVm
                {
                    MsgInfo = @"Sincroniza Usuario entre el AD y el O365",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.Sincronizar.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(logSync);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var exitoProceso = false;
                var exitoGuardar = HomeSysWebFactory.CrearUsuario(model);
                if (exitoGuardar)
                {
                    //Obtener Codigo Licencia                    
                    var mantLicObjVm = MantenedorLicenciaFactory.ObtenerLicenciaCodigo(model.CodigoLicencia.Trim());
                    decimal licenciaId = mantLicObjVm.LicenciaId;

                    //Ingresa datos usuarios a base de datos
                    var estadoUsr = new EstadoCuentaUsuarioVm {
                        Apellidos = model.Apellidos,
                        CodigoLicencia = !string.IsNullOrEmpty(model.CodigoLicencia.Trim()) ? model.CodigoLicencia.Trim() : string.Empty,
                        Correo = model.Correo,
                        CreadoAd = true,
                        CuentaAd = model.NombreUsuario,
                        Descripcion = model.Descripcion,
                        Dominio = model.UpnPrefijo,
                        Eliminado = false,
                        FechaCreacion = DateTime.Now,
                        Habilitado = true,
                        LicenciaId = licenciaId, 
                        Nombres = model.Nombres,
                        Sincronizado = false,
                        Clave = model.Clave
                    };

                    try
                    {
                        EstadoCuentaUsuarioFactory = new EstadoCuentaUsuarioFactory();
                        EstadoCuentaUsuarioFactory.CrearEstadoCuentaUsuarioDirecto(estadoUsr);
                        //HiloEstadoCuentaUsuario.CrearEstadoCuentaUsuario(estadoUsr);                            
                        exitoProceso = true;
                    }
                    catch (Exception ex)
                    {
                        Utils.LogErrores(ex);
                    }                    
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = exitoProceso,
                        Error = exitoProceso == false ? "Proceso de guardar y sincronizar terminado incorrectamente, favor intentar más tarde." : string.Empty,
                        Session = varSession,
                        CodigoRuta = codRuta
                    }
                };
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = ex.Message,
                        Session = varSession,
                        CodigoRuta = codRuta
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
                        Validar = true,
                        Error = ex.Message,
                        Session = varSession,
                        CodigoRuta = codRuta
                    }
                };
            }
        }

        [HttpPost]
        public ActionResult ActualizarCuentaUsuario(HomeSysWebVm model)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {                    
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                if (string.IsNullOrEmpty(model.Nombres)) throw new ArgumentException("Ingresar nombres del usuario");
                if (string.IsNullOrEmpty(model.Apellidos)) throw new ArgumentException("Ingresar apellidos del usuario");
                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");
                if (string.IsNullOrEmpty(model.Correo)) throw new ArgumentException("Asignar correo");
                if (string.IsNullOrEmpty(model.UpnPrefijo)) throw new ArgumentException("Ingresar Dominio");
                if (string.IsNullOrEmpty(model.Descripcion)) throw new ArgumentException("Ingresar Descripción");
                if (string.IsNullOrEmpty(model.CodigoLicencia)) throw new ArgumentException("Ingresar Código Licencia");
                //if (string.IsNullOrEmpty(model.Clave)) throw new ArgumentException("Ingresar clave del usuario");                                
                if (string.IsNullOrEmpty(model.Clave)) model.Clave = @"***";
                if (string.IsNullOrEmpty(model.PatchOu))
                {
                   throw new ArgumentException("Seleccionar unidad organizativa");                   
                }
                
                //Validar que el código ingresado exista
                if (!MantenedorLicenciaFactory.ExisteCodigoLicencia(model.CodigoLicencia.Trim()))
                {
                    throw new ArgumentException("El código licencia ingresado no existe en la base de datos.");
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var exitoProceso = HomeSysWebFactory.ActualizarCuentaUsuario(model);

                if (exitoProceso)
                {
                    //Obtener Codigo Licencia                    
                    var mantLicObjVm = MantenedorLicenciaFactory.ObtenerLicenciaCodigo(model.CodigoLicencia.Trim());
                    decimal licenciaId = mantLicObjVm.LicenciaId;

                    //Ingresa datos usuarios a base de datos
                    var estadoUsr = EstadoCuentaUsuarioFactory.GetObjetoEstadoCuentaUsuarioAllByCuenta(model.NombreUsuario.Trim());
                    if (estadoUsr != null)
                    {
                        try
                        {
                            //Valida que usuario exista en Portal
                            var ctaSync = HomeSysWebFactory.ExisteUsuarioPortal(model.Correo.Trim());
                            var ctaLic = false;
                            if (ctaSync)
                            {
                                ctaLic = HomeSysWebFactory.ExisteLicenciaUsuarioPortal(model.Correo.Trim());
                            }
                            estadoUsr.Apellidos = model.Apellidos.Trim();
                            estadoUsr.Clave = model.Clave.Trim();
                            estadoUsr.CodigoLicencia = model.CodigoLicencia.Trim();
                            estadoUsr.Correo = model.Correo.Trim();
                            estadoUsr.Descripcion = model.Descripcion.Trim();
                            estadoUsr.Dominio = model.UpnPrefijo.Trim();                            
                            estadoUsr.LicenciaId = licenciaId;
                            estadoUsr.Nombres = model.Nombres.Trim();                            
                            estadoUsr.Eliminado = false;
                            estadoUsr.FechaBaja = null;
                            estadoUsr.Habilitado = true;
                            estadoUsr.Vigente = true;
                            estadoUsr.Sincronizado = ctaSync;
                            estadoUsr.LicenciaAsignada = ctaLic;

                            EstadoCuentaUsuarioFactory.ActualizaEstadoCuentaUsuario(estadoUsr);
                        }
                        catch (Exception ex)
                        {
                            Utils.LogErrores(ex);
                        }
                    }
                    else {
                        //Valida que usuario exista en Portal
                        var ctaSync = HomeSysWebFactory.ExisteUsuarioPortal(model.Correo.Trim());
                        var ctaLic = false;
                        if (ctaSync)
                        {
                            ctaLic = HomeSysWebFactory.ExisteLicenciaUsuarioPortal(model.Correo.Trim());
                        }
                        //Ingresa datos usuarios a base de datos
                        estadoUsr = new EstadoCuentaUsuarioVm
                        {
                            Apellidos = model.Apellidos,
                            CodigoLicencia = model.CodigoLicencia.Trim(),
                            Correo = model.Correo,
                            CreadoAd = true,
                            CuentaAd = model.NombreUsuario.Trim(),
                            Descripcion = model.Descripcion.Trim(),
                            Dominio = model.UpnPrefijo,
                            Eliminado = false,
                            FechaCreacion = DateTime.Now,                           
                            Habilitado = true,
                            LicenciaId = licenciaId,
                            Nombres = model.Nombres,
                            Sincronizado = ctaSync,
                            Clave = model.Clave,
                            Vigente = true,
                            LicenciaAsignada = ctaLic
                        };

                        try
                        {
                            EstadoCuentaUsuarioFactory = new EstadoCuentaUsuarioFactory();
                            EstadoCuentaUsuarioFactory.CrearEstadoCuentaUsuarioDirecto(estadoUsr);                                                        
                        }
                        catch (Exception ex)
                        {
                            Utils.LogErrores(ex);
                        }
                    }                    
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = exitoProceso,
                        Error = exitoProceso == false ? "Proceso de actualizar cuenta usuario terminado incorrectamente, favor intentar más tarde." : string.Empty,
                        Session = varSession
                    }
                };
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = ex.Message,
                        Session = varSession
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
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        /// <summary>
        /// Sincronizar Usuarios
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SincronizarUsuario(HomeSysWebVm model)
        {
            try
            {
                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");
                if (string.IsNullOrEmpty(model.UpnPrefijo)) throw new ArgumentException("Ingresar Dominio");

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Sincroniza Usuario entre el AD y el O365",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.Sincronizar.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var ejecSync = HomeSysWebFactory.ForzarDirSync(model);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = ejecSync,
                        Error = string.Empty
                    }
                };
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = ex.Message
                    }
                };
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al sincronizar usuario. Si el problema persiste contacte a soporte IT" });
            }
        }
        /// <summary>
        /// Asignar Licencia
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AsignarLicenciaUsuario(HomeSysWebVm model)
        {
            try
            {
                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");
                if (string.IsNullOrEmpty(model.UpnPrefijo)) throw new ArgumentException("Ingresar Dominio");
                if (string.IsNullOrEmpty(model.Licencia)) throw new ArgumentException("Ingresar licencia del usuario");

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Asigna Licencia Usuario en el O365",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.AsignarLicencia.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }
                
                HomeSysWebFactory = new HomeSysWebFactory();
                var ejecAsig = HomeSysWebFactory.AsignarLicenciaUsuario(model);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = ejecAsig,
                        Error = string.Empty
                    }
                };
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Error = ex.Message
                    }
                };
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al asignar licencia usuario. Si el problema persiste contacte a soporte IT" });
            }
        }

        [HttpPost]
        public ActionResult DeshabilitarUsuario(HomeSysWebVm model)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Deshabilita Usuario en el AD",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.DeshabilitarUsuario.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    //HiloEstadoSincronizacion.ActualizarEstadoSync(false, SessionViewModel.Usuario.Nombre.Trim(), "D");
                    Utils.LogErrores(ex);
                }

                //HiloEstadoSincronizacion.ActualizarEstadoSync(true, SessionViewModel.Usuario.Nombre.Trim(), "D");
                HomeSysWebFactory = new HomeSysWebFactory();
                var username = model.NombreUsuario.ToLower().Trim();
                var procExito = HomeSysWebFactory.DeshabilitarUsuarioAd(username);
                //if (procExito)
                //{
                //    procExito = HomeSysWebFactory.DeshabilitarComputadorAd(username);
                //}                

                if (procExito)
                {
                    //Actualiza cuenta usuario BD cambio Habilitado en False
                    var estaCtaUsrUnico = EstadoCuentaUsuarioFactory.GetObjetoEstadoCuentaUsuarioByCuenta(model.NombreUsuario.Trim());
                    if (estaCtaUsrUnico != null)
                    {
                        estaCtaUsrUnico.Habilitado = false;
                        estaCtaUsrUnico.Sincronizado = false;
                        estaCtaUsrUnico.LicenciaAsignada = false;
                        estaCtaUsrUnico.FechaBaja = DateTime.Now;
                        EstadoCuentaUsuarioFactory.ActualizaEstadoCuentaUsuario(estaCtaUsrUnico);
                    }
                    else
                    {
                        decimal licenciaId = 1;
                        var codigoLicencia = string.Empty;
                        if (!string.IsNullOrEmpty(model.CodigoLicencia))
                        {
                            var mantLicObjVm = MantenedorLicenciaFactory.ObtenerLicenciaCodigo(model.CodigoLicencia);
                            if (mantLicObjVm != null)
                            {
                                licenciaId = mantLicObjVm.LicenciaId;
                            }
                        }
                        else
                        {
                            var objLic = MantenedorLicenciaFactory.GetMantenedorLicencia(licenciaId);
                            if (objLic != null)
                            {
                                codigoLicencia = objLic.Codigo;
                            }
                        }
                        //Ingresa datos usuarios a base de datos
                        var estaCtaUsrNew = new EstadoCuentaUsuarioVm
                        {
                            Apellidos = model.Apellidos,
                            CodigoLicencia = string.IsNullOrEmpty(model.CodigoLicencia) ? codigoLicencia : model.CodigoLicencia.Trim(),
                            Correo = model.Correo.Trim(),
                            CreadoAd = true,
                            CuentaAd = model.NombreUsuario.Trim(),
                            Descripcion = model.Descripcion.Trim(),
                            Dominio = model.UpnPrefijo,
                            Eliminado = false,
                            FechaCreacion = DateTime.Now,
                            Habilitado = false,
                            LicenciaId = licenciaId,
                            Nombres = model.Nombres.Trim(),
                            Sincronizado = false,
                            Clave = string.IsNullOrEmpty(model.Clave) ? @"Inicio01" : model.Clave.Trim(),
                            Vigente = true,
                            FechaBaja = DateTime.Now,
                            LicenciaAsignada = false                            
                        };

                        try
                        {
                            EstadoCuentaUsuarioFactory = new EstadoCuentaUsuarioFactory();
                            EstadoCuentaUsuarioFactory.CrearEstadoCuentaUsuarioDirecto(estaCtaUsrNew);
                        }
                        catch (Exception ex)
                        {
                            //HiloEstadoSincronizacion.ActualizarEstadoSync(false, SessionViewModel.Usuario.Nombre.Trim(), "D");
                            Utils.LogErrores(ex);
                        }
                    }

                    //var usuarioModificacion = SessionViewModel.Usuario.Nombre.Trim();
                    //var listaEstCuentaVmHilo = new List<object>
                    //{                        
                    //    usuarioModificacion
                    //};                    
                    //HomeSysWebFactory.ForzarDirSync();
                    //HiloEstadoSincronizacion.ActualizarEstadoSync(false, SessionViewModel.Usuario.Nombre.Trim(), "D");

                    //Ejecuta Hilo para el proceso de sync de cuentas
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();                   

                    _hiloEjecucion = new Thread(InciarProcesoHiloSincronizarCuenta)
                    {
                        IsBackground = true,
                        Priority = ThreadPriority.Highest
                    };
                    _hiloEjecucion.Start();
                    //_hiloEjecucion.Start(listaEstCuentaVmHilo);
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = procExito,
                        Error = string.Empty,
                        Session = varSession
                    }
                };
            }
            catch (Exception ex)
            {
                //HiloEstadoSincronizacion.ActualizarEstadoSync(false, SessionViewModel.Usuario.Nombre.Trim(), "D");
                Utils.LogErrores(ex);
                return new JsonResult
                {
                    Data = new
                    {
                        Validar = false,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [SoapDocumentMethod(OneWay = true)]
        public void InciarProcesoHiloSincronizarCuenta()
        {            
            try
            {
                //HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "D");
                HomeSysWebFactory = new HomeSysWebFactory();
                try
                {
                    HomeSysWebFactory.ForzarDirSync();
                }
                catch (Exception ex)
                {
                    //HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "S");
                    Utils.LogErrores(ex);
                }
            }
            catch (Exception ex)
            {
                //HiloEstadoSincronizacion.ActualizarEstadoSync(false, usuarioModificacion, "D");
                var msgError = @"Error en proceso asincronico Syncronizar por cuenta : " + ex.Message;
                var exNew = new Exception(msgError);
                Utils.LogErrores(exNew);
            }
        }

        /// <summary>
        /// Buscar el grupo ingresado en el AD 
        /// </summary>
        /// <param name="nomGrupo"></param>
        /// <returns></returns>

        public JsonResult ObtenerGrupoAd(string nomGrupo)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Grupo = (List<GrupoAdVm>)null,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Asignar Grupos",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.AsignarGrupo.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }

                var listaGrupos = HomeSysWebFactory.ObtenerGrupoAdByNombre(nomGrupo);

                return new JsonResult
                {
                    Data = new
                    {
                        Grupo = listaGrupos,
                        Error = string.Empty,
                        Session = varSession
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
                        Grupo = (List<GrupoAdVm>) null,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [HttpPost]
        public JsonResult ObtenerListaGruposAdOu(string patchOu)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Grupo = (List<GrupoAdVm>)null,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var listaGrupos = HomeSysWebFactory.ObtenerListadoGrupoAdByOu(patchOu);

                return new JsonResult
                {
                    Data = new
                    {
                        GrupoOu = listaGrupos,
                        Error = string.Empty,
                        Session = varSession
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
                        GrupoOu = (List<GrupoAdVm>)null,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [HttpPost]
        public JsonResult ObtenerListaGruposAdOu2(string patchOu, string usrAD)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Grupo = (List<GrupoAdVm>)null,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();
                var listaGrupos = HomeSysWebFactory.ObtenerListadoGrupoAdByOu(patchOu, usrAD);

                return new JsonResult
                {
                    Data = new
                    {
                        GrupoOu = listaGrupos,
                        Error = string.Empty,
                        Session = varSession
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
                        GrupoOu = (List<GrupoAdVm>)null,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [HttpPost]
        public ActionResult GuardarGrupos(string nomGrupos, string userName, string upnPrefijo)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            TieneLicencia = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();

                //Seccion registra Log Transaccional
                var log = new LogInfoVm
                {
                    MsgInfo = @"Guardar Grupos",
                    UserInfo = SessionViewModel.Usuario.Nombre,
                    FechaInfo = DateTime.Now,
                    AccionIdInfo = EnumAccionInfo.GuardarGrupo.GetHashCode()
                };

                try
                {
                    LogInfoFactory.CrearLogInfo(log);
                }
                catch (Exception ex)
                {
                    Utils.LogErrores(ex);
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = HomeSysWebFactory.GuardarGrupos(nomGrupos, userName),
                        TieneLicencia = HomeSysWebFactory.ExisteLicenciaUsuarioPortal(userName + upnPrefijo),
                        Error = string.Empty,
                        Session = varSession
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
                        TieneLicencia = false,
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [HttpPost]
        public ActionResult AsociarGrupoUsuario(string nomGrupo, string userName)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = HomeSysWebFactory.AsociarGrupoUsuario(nomGrupo, userName),
                        Error = string.Empty,
                        Session = varSession
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
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }

        [HttpPost]
        public ActionResult DesAsociarGrupoUsuario(string nomGrupo, string userName)
        {
            var varSession = true;
            try
            {
                //Verifica que sesiones no sean nulas, si lo es redirecciona a página login                
                if (System.Web.HttpContext.Current.Session["UsuarioVM"] == null || System.Web.HttpContext.Current.Session["EstructuraArbol"] == null)
                {
                    varSession = false;
                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = false,
                            Error = string.Empty,
                            Session = varSession
                        }
                    };
                }

                HomeSysWebFactory = new HomeSysWebFactory();

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = HomeSysWebFactory.DesAsociarGrupoUsuario(nomGrupo, userName),
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
                        Error = ex.Message,
                        Session = varSession
                    }
                };
            }
        }
    }
}
