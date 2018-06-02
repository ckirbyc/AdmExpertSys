using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.Mapping.Factories;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Data.Entity.Infrastructure;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class UsuarioPerfilController : BaseController
    {
        // GET: UsuarioPerfil
        protected UsuarioFactory UsuarioFactory;        

        public UsuarioPerfilController()
        {
        }

        public UsuarioPerfilController(UsuarioFactory usuarioFactory)
        {
            UsuarioFactory = usuarioFactory;
        }

        public ActionResult Index()
        {
            try
            {
                //Obtener todos los usuarios del sistema
                var listaUsrPerfil = UsuarioFactory.GetAllUsuarioPerfil();
                return View(listaUsrPerfil);
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al cargar página Usuario Perfil. Si el problema persiste contacte a soporte IT" });
            }
        }

        public ActionResult Create()
        {
            return View(UsuarioFactory.GetCreateView());
        }

        [HttpPost]
        public ActionResult GuardarUsuario(string usuario, string perfil, string esAdm)
        {
            try
            {
                var objVm = new UsuarioPerfilVm {
                    Usuario1 = usuario.ToLower().Trim(),
                    PerfilId = Convert.ToInt32(perfil),
                    EsAdm = Convert.ToBoolean(esAdm),
                    Visible = true
                };

                if (!UsuarioFactory.ValidarExistenciaUsuarioBd(objVm.Usuario1))
                {
                    var exito = UsuarioFactory.CrearUsuarioPerfil(objVm);

                    return new JsonResult
                    {
                        Data = new
                        {
                            Validar = exito,
                            Mensaje = string.Empty
                        }
                    };
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Validar = true,
                        Mensaje = @"No se puede guardar el usuario porque este ya existe. Favor vuelva a intentar."
                    }
                };

            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error", new { message = "Error al guardar usuario. Si el problema persiste contacte a soporte IT" });
            }            
        }

        public ActionResult Edit(decimal id)
        {
            return View(UsuarioFactory.GetUsuarioPerfil(id));
        }

        [HttpPost]
        public ActionResult Edit(UsuarioPerfilVm model)
        {
            try
            {
                if (model == null) throw new ArgumentNullException("model");

                if (!ModelState.IsValid)
                {
                    model = UsuarioFactory.GetCreateView();
                    return View(model);
                }
                model.Visible = true;
                UsuarioFactory.ActualizaPerfilUsuario(model);

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Utils.LogErrores(ex);
                return RedirectToAction("Index", "Error",
                    new { message = "Error al intentar guardar registro de Usuario Perfil. Por favor reinténtelo más tarde." });
            }
        }

        public ActionResult Delete(decimal id)
        {
            return View(UsuarioFactory.GetUsuarioPerfil(id));
        }

        [HttpPost]
        public ActionResult Delete(UsuarioPerfilVm model)
        {
            try
            {
                var objVm = UsuarioFactory.GetUsuarioPerfil(model.Id);
                UsuarioFactory.EliminaPerfilUsuario(objVm);

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
                    new { message = "Error al intentar eliminar registro de Perfil Usuario. Por favor reinténtelo más tarde." });
            }
        }
    }
}