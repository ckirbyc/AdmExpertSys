using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Core.Domain.Enums;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.Models;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class HomeController : BaseController
    {
        protected HomeSysWebFactory HomeSysWebFactory;
        protected LogInfoFactory LogInfoFactory;

        public HomeController()
        {
        }

        public HomeController(LogInfoFactory logInfoFactory)
        {
            LogInfoFactory = logInfoFactory;
        }

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
                    var estructura = HomeSysWebFactory.GetArquitecturaArbolAd();
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
        /// <summary>
        /// Controler valida credenciales del usuario al ingresar al sistema.
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ValidaLogin(LoginVm model)
        {
            try
            {
                System.Web.HttpContext.Current.Session["UsuarioVM"] = "asd";
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
                    model.EstaAutenticado = true;
                    return RedirectToAction("HomeSysWeb", "Home");
                }

                model.EstaAutenticado = false;
                return RedirectToAction("IndexLogin", "Error", new { message = "Error, usuario o contraseña no corresponden. Si el problema persiste contacte a soporte IT" });
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("IndexLogin", "Error", new { message = "Error al iniciar sesión. Si el problema persiste contacte a soporte IT" });
            }
        }

        [HttpPost]
        public ActionResult VerificarUsuario(string nombreUsuario)
        {
            try
            {
                if (string.IsNullOrEmpty(nombreUsuario)) throw new ArgumentNullException("nombreUsuario");

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
                //var sUpn = nombreUsuario.Trim() + "@agrosuper.com";
                //var existe = HomeSysWebFactory.ExisteLicenciaUsuarioPortal(sUpn);
                var usuarioAd = HomeSysWebFactory.ObtenerUsuarioExistente(nombreUsuario.Trim());
                bool chequear = usuarioAd != null;

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = chequear,
                        DatosUsuario = usuarioAd
                    }
                };
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al validar usuario. Si el problema persiste contacte a soporte IT" });
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
            try
            {
                if (string.IsNullOrEmpty(model.Nombres)) throw new ArgumentException("Ingresar nombres del usuario");
                if (string.IsNullOrEmpty(model.Apellidos)) throw new ArgumentException("Ingresar apellidos del usuario");
                if (string.IsNullOrEmpty(model.NombreUsuario)) throw new ArgumentException("Ingresar ID del usuario");
                if (string.IsNullOrEmpty(model.UpnPrefijo)) throw new ArgumentException("Ingresar Dominio");
                if (string.IsNullOrEmpty(model.Descripcion)) throw new ArgumentException("Ingresar Descripción");
                if (model.ExisteUsuario == false)
                {
                    if (string.IsNullOrEmpty(model.Clave)) throw new ArgumentException("Ingresar clave del usuario");
                }
                else
                {
                    model.Clave = @"$aaa123";
                }
                if (string.IsNullOrEmpty(model.PatchOu))
                {
                    if (model.ExisteUsuario == false)
                    {
                        throw new ArgumentException("Seleccionar unidad organizativa");
                    }
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
                    //Task.Delay(TimeSpan.FromSeconds(10)).Wait();
                    var ejecSync = HomeSysWebFactory.ForzarDirSync(model);
                    if (ejecSync)
                    {
                        exitoProceso = true;
                    }
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = exitoProceso,
                        Error = exitoProceso == false ? "Proceso de guardar y sincronizar terminado incorrectamente, favor intentar más tarde." : string.Empty
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
                    new {message = "Error al ingresar usuario. Si el problema persiste contacte a soporte IT"});
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
            try
            {
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
                    Utils.LogErrores(ex);
                }
                
                HomeSysWebFactory = new HomeSysWebFactory();
                var username = model.NombreUsuario.ToLower().Trim();
                var procExito = HomeSysWebFactory.DeshabilitarUsuarioAd(username);

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = procExito,
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
        /// <summary>
        /// Buscar el grupo ingresado en el AD 
        /// </summary>
        /// <param name="nomGrupo"></param>
        /// <returns></returns>
        
        public JsonResult ObtenerGrupoAd(string nomGrupo)
        {
            try
            {
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
                        Grupo = (List<GrupoAdVm>) null,
                        Error = ex.Message
                    }
                };
            }
        }

        [HttpPost]
        public ActionResult GuardarGrupos(string nomGrupos, string userName, string upnPrefijo)
        {
            try
            {
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
                        TieneLicencia = false,
                        Error = ex.Message
                    }
                };
            }
        }
    }
}
