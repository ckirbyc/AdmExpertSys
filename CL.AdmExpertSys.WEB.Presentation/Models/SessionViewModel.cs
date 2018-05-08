
using CL.AdmExpertSys.WEB.Core.Domain.Dto;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;
using System;
using System.Web;

namespace CL.AdmExpertSys.WEB.Presentation.Models
{
    [Serializable]
    public static class SessionViewModel
    {
        /// <summary>
        /// Nombre usuario Completo para Mostrar en Formularios
        /// </summary>
        public static UsuarioVm Usuario
        {
            get
            {
                return (UsuarioVm)HttpContext.Current.Session["UsuarioVM"];
            }
            set
            {
                HttpContext.Current.Session["UsuarioVM"] = value;
            }
        }

        public static EstructuraArbolAd EstructuraArbolAd
        {
            get
            {
                return (EstructuraArbolAd)HttpContext.Current.Session["EstructuraArbol"];
            }
            set
            {
                HttpContext.Current.Session["EstructuraArbol"] = value;
            }
        }
    }
}