using System.Collections.Generic;
using System.IO;
using System.Reflection;
using CL.AdmExpertSys.Web.Infrastructure.Helpers;
using CL.AdmExpertSys.WEB.Presentation.Models;
using System;
using System.Globalization;
using System.Threading;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Controllers
{
    [HandleError]
    public class BaseController : Controller
    {
        /// <summary>
        /// Inicializa todos los controllers.
        /// </summary>
        /// <param name="requestContext"></param>
        protected override void Initialize(System.Web.Routing.RequestContext requestContext)
        {
            base.Initialize(requestContext);

            var cultureInfo = CultureInfo.GetCultureInfo("es-CL");

            Thread.CurrentThread.CurrentCulture = cultureInfo;
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
        }

        /// <summary>
        /// Envía un html en String.
        /// </summary>
        /// <returns></returns>
        protected string RenderPartialViewToString()
        {
            return RenderPartialViewToString(null, null);
        }

        /// <summary>
        /// Envía un html en String recibiendo un viewName.
        /// </summary>
        /// <param name="viewName"></param>
        /// <returns></returns>
        protected string RenderPartialViewToString(string viewName)
        {
            return RenderPartialViewToString(viewName, null);
        }

        /// <summary>
        /// Envía un html en String recibiendo un model.
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        protected string RenderPartialViewToString(object model)
        {
            return RenderPartialViewToString(null, model);
        }

        /// <summary>
        /// Envía un html en String recibiendo un viewName y un model.
        /// </summary>
        /// <param name="viewName"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        protected string RenderPartialViewToString(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
                viewName = ControllerContext.RouteData.GetRequiredString("action");

            ViewData.Model = model;

            using (var sw = new StringWriter())
            {
                var viewResult = ViewEngines.Engines.FindPartialView(ControllerContext, viewName);
                var viewContext = new ViewContext(ControllerContext, viewResult.View, ViewData, TempData, sw);
                viewResult.View.Render(viewContext, sw);

                return sw.GetStringBuilder().ToString();
            }
        }

        /// <summary>
        /// El manejador de los métodos de los controladores.
        /// </summary>
        /// <param name="filterContext"></param>
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var action = (string)filterContext.RouteData.Values["action"];

            MethodInfo[] metodosControlador = filterContext.Controller.GetType().GetMethods();

            var parametrosControlador = new Dictionary<string, Type>();

            foreach (var item in filterContext.ActionParameters)
            {
                parametrosControlador.Add(item.Key, item.Value == null ? null : item.Value.GetType());
            }


            MethodInfo metodoEncontrado = null;

            foreach (MethodInfo item in metodosControlador)
            {
                bool parametrosCoinciden = true;
                if (item.Name == action)
                {
                    ParameterInfo[] parametros = item.GetParameters();

                    foreach (ParameterInfo itemParametro in parametros)
                    {
                        if (!parametrosControlador.ContainsKey(itemParametro.Name))
                        {
                            parametrosCoinciden = false;
                            break;
                        }
                    }

                    if (parametrosCoinciden)
                        metodoEncontrado = item;
                    else
                        continue;
                }
            }


            bool permiteVerSinAutenticacion = (metodoEncontrado.GetCustomAttributes(typeof(AllowNotAutenticate), true).Length > 0);

            if (permiteVerSinAutenticacion)
            {
                base.OnActionExecuting(filterContext);
                return;
            }

            System.Web.HttpContext ctx = System.Web.HttpContext.Current;

            // check  sessions here
            if (ctx.Session["UsuarioVM"] == null || ctx.Session["EstructuraArbol"] == null)
            {
                filterContext.Result = RedirectToAction("IndexLogin", "Error", new { mensajeError = "Su sesión ha expirado." });
                return;
            }            

            base.OnActionExecuting(filterContext);
        }

        /// <summary>
        /// Realiza un ParseEnum genérico sobre el parámetro de entrada.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T ParseEnum<T>(string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }

        public static string GetUsuarioRut()
        {
            if (SessionViewModel.Usuario.Rut > 0)
            {
                return SessionViewModel.Usuario.Rut + "-" + SessionViewModel.Usuario.Dv;
            }
            return SessionViewModel.Usuario.Nombre;
        }
    }
}
