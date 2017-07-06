
using System;
using System.Collections.Generic;

namespace CL.AdmExpertSys.Web.Infrastructure.LogTransaccional
{
    public class LogMessage
    {
        #region Propiedades

        /// <summary>
        /// Metodo o Procedimiento desde el cual se ha generado el log
        /// </summary>
        public string Modulo { get; set; }

        /// <summary>
        /// Corresponde a la funcionalidad realizada (CRUD) sobre una entidad de negocio.
        /// </summary>
        public string Funcionalidad { get; set; }

        /// <summary>
        /// Indica la fecha y hora en que el log se ha ingresado al aplicativo
        /// </summary>
        public DateTime FechaRegistro { get; set; }

        /// <summary>
        /// La IP asociada al cliente conectado
        /// </summary>
        /// <example>192.168.0.1</example>
        public string IpCliente { get; set; }

        /// <summary>
        /// Servidor en el cual se ha generado el log
        /// </summary>
        public string Servidor { get; set; }

        /// <summary>
        /// Nombre y Versión del navegador web asociado al cliente conectado
        /// </summary>
        public string DetalleBrowser { get; set; }

        /// <summary>
        /// RUT del usuario interno que ha realizado consultas
        /// </summary>
        public int RutUsuario { get; set; }

        /// <summary>
        /// DV del rut de usuario interno que ha realizado consultas
        /// </summary>
        public char DvRutUsuario { get; set; }


        /// <summary>
        /// Nombre de Usuario autenticado 
        /// </summary>
        public string NombreUsuario { get; set; }

        /// <summary>
        /// Objeto que contiene el grafo perteneciente a la entrada de llamado a un servicio
        /// </summary>
        public Dictionary<string, object> Entrada { get; set; }

        /// <summary>
        /// Objeto que contiene el grafo perteneciente a la salida de algun servicio
        /// </summary>
        public object Salida { get; set; }

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        public LogMessage()
        {
            Entrada = new Dictionary<string, object>();
        }
    }
}
