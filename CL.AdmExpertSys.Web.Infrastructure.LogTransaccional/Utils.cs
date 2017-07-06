using CL.AdmExpertSys.Web.Infrastructure.Helpers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml;

namespace CL.AdmExpertSys.Web.Infrastructure.LogTransaccional
{
    /// <summary>
    /// Utilitarios Generales
    /// </summary>
    /// <remarks>
    /// $Revision: 25749 $
    /// $Author: ckirby $
    /// $Date: 2016-05-16 17:51:19 -0300 (Tue, 27 May 2016) $
    /// </remarks>
    public static class Utils
    {
        /// <summary>
        /// Obtiene la parte entera de un rut en el formato XX.XXX.XXX-X
        /// </summary>
        /// <returns></returns>
        public static int ObtenerParteEnteraDeRut(string rutCompleto)
        {
            if (String.IsNullOrEmpty(rutCompleto))
            {
                return 0;
            }

            if (rutCompleto.IndexOf('.') > 0)
            {
                rutCompleto = rutCompleto.Replace(".", String.Empty);
            }

            if (rutCompleto.IndexOf('-') > 0)
            {
                rutCompleto = rutCompleto.Substring(0, rutCompleto.IndexOf('-'));
            }

            int rutEntero;
            Int32.TryParse(rutCompleto, out rutEntero);
            return rutEntero;
        }

        /// <summary>
        /// Método que calcula el dígito verificador de un rut
        /// </summary>
        /// <param name="rut">Parte entera del Rut a la cual se desea calcular su dígito verificador</param>
        /// <returns>String con el dígito verificador, que puede ser un dígito o la letra 'K'</returns>
        public static string CalcularDigitoVerificador(int rut)
        {
            int contador = 2;
            int acumulador = 0;
            while (rut != 0)
            {
                int multiplo = (rut % 10) * contador;
                acumulador = acumulador + multiplo;
                rut = rut / 10;
                contador = contador + 1;
                if (contador == 8)
                {
                    contador = 2;
                }
            }

            int digito = 11 - (acumulador % 11);
            string rutDigito = digito.ToString(CultureInfo.InvariantCulture).Trim();
            if (digito == 10)
            {
                rutDigito = "K";
            }
            if (digito == 11)
            {
                rutDigito = "0";
            }
            return (rutDigito);
        }

        /// <summary>
        /// Obtiene Dv formateado
        /// </summary>
        /// <param name="dv"></param>
        /// <returns></returns>
        public static byte[] ObtenerDv(string dv)
        {
            if (!String.IsNullOrEmpty(dv))
                return Encoding.UTF8.GetBytes(dv);
            return null;
        }

        /// <summary>
        /// Convierte un datatable a formato CSV
        /// </summary>
        /// <param name="table"></param>
        /// <param name="delimiter"></param>
        /// <param name="includeHeader"></param>
        /// <returns></returns>
        public static string ToCsv(DataTable table, string delimiter, bool includeHeader)
        {
            var result = new StringBuilder();

            if (includeHeader)
            {
                foreach (DataColumn column in table.Columns)
                {
                    result.Append(column.ColumnName);
                    result.Append(delimiter);
                }

                result.Remove(--result.Length, 0);
                result.Append(Environment.NewLine);
            }

            foreach (DataRow row in table.Rows)
            {
                foreach (object item in row.ItemArray)
                {
                    if (item is DBNull)
                        result.Append(delimiter);
                    else
                    {
                        string itemAsString = item.ToString();
                        itemAsString = itemAsString.Replace("\"", "\"\"");
                        itemAsString = "\"" + itemAsString + "\"";
                        result.Append(itemAsString + delimiter);
                    }
                }

                result.Remove(--result.Length, 0);
                result.Append(Environment.NewLine);
            }
            return result.ToString();
        }

        /// <summary>
        /// Obtiene el tamañana en bytes del objeto de entrada
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int GetObjectSize(object value)
        {
            var bf = new BinaryFormatter();
            using (var ms = new MemoryStream())
            {
                bf.Serialize(ms, value);
                var array = ms.ToArray();
                return array.Length;
            }
        }

        /// <summary>
        /// Convierte un objeto que representa la salida de la llamada a un servicio transaccional
        /// </summary>
        /// <param name="objetoEntrada">El objeto transaccional retornado</param>
        /// <param name="profundidad">La cantidad de niveles (clases internas ) que se quiere convertir a string</param>
        /// <returns></returns>
        public static string ConvertObjetoEntradaToString(Dictionary<string, object> objetoEntrada, int profundidad)
        {
            return objetoEntrada.Aggregate(String.Empty, (current, kvp) => current + (kvp.Key + " : \n" + ConvertObjetoSalidaToString(kvp.Value, profundidad)["salida"] + "\n"));
        }

        /// <summary>
        /// Convierte un objeto que representa la salida de la llamada a un servicio transaccional
        /// </summary>
        /// <param name="objetoSalida">El objeto transaccional retornado</param>
        /// <param name="profundidad">La cantidad de niveles (clases internas ) que se quiere convertir a string</param>
        /// <returns></returns>
        public static Dictionary<string, string> ConvertObjetoSalidaToString(object objetoSalida, int profundidad)
        {
            var diccSalida = new Dictionary<string, string>();
            try
            {
                var sb = new StringBuilder();

                DataTable dt;
                if (objetoSalida is DataSet)
                {
                    dt = (objetoSalida as DataSet).Tables[0];
                    diccSalida["salida"] = ToCsv(dt, " ", true);
                    return diccSalida;
                }

                if (objetoSalida is DataTable)
                {
                    dt = objetoSalida as DataTable;
                    diccSalida["salida"] = ToCsv(dt, " ", true);
                    return diccSalida;
                }

                using (var sw = new StringWriter(sb, CultureInfo.InvariantCulture))
                {
                    ObjectDumper.Write(objetoSalida, profundidad, sw);
                }

                diccSalida = LimitarLargoEnBytes(sb.ToString(), Convert.ToInt32(ConfigurationManager.AppSettings["MaxByteSizeSalida"], CultureInfo.InvariantCulture));
                return diccSalida;
            }
            catch (InvalidCastException)
            {
                // logger.Errores(String.Format(CultureInfo.InvariantCulture, "Error al convertir objeto, no es compatible.", ex.Message, ex.InnerException));
            }

            diccSalida["comentario"] = String.Empty;
            return diccSalida;
        }

        /// <summary>
        /// Limitar el largo de un objeto a x bytes
        /// </summary>
        /// <param name="input"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        public static Dictionary<string, string> LimitarLargoEnBytes(string input, int maxLength)
        {
            var comentarioSalida = String.Empty;
            var nuevaSalida = new string(input
                                             .TakeWhile((c, i) =>
                                                        Encoding.UTF8.GetByteCount(input.Substring(0, i + 1)) <= maxLength)
                                             .ToArray());

            if (nuevaSalida.Length < input.Length)
            {
                comentarioSalida = "Se limitó el tamaño de la salida.";
            }

            var diccSalida = new Dictionary<string, string>();
            diccSalida["salida"] = nuevaSalida;
            diccSalida["comentario"] = comentarioSalida;
            return diccSalida;
        }

        public static void LogErrores(Exception exception)
        {
            try
            {
                var dirLog = ConfigurationManager.AppSettings["UrlLog"];
                var exceptionSerializeError = new ExceptionSerialize(exception);
                string xmlTexto = Serializacion.XmlSerialize(exceptionSerializeError);

                var doc = new XmlDocument { XmlResolver = null };
                doc.LoadXml(xmlTexto);

                XmlNode root = doc.DocumentElement;

                //Create a new node.
                XmlElement elemSource = doc.CreateElement("Source");
                elemSource.InnerText = exception.Source;
                root.AppendChild(elemSource);

                //Create a new node.
                XmlElement elemMessage = doc.CreateElement("Message");
                elemMessage.InnerText = exception.Message;
                root.AppendChild(elemMessage);

                if (exception.InnerException != null)
                {
                    //Create a new node.
                    XmlElement elemInner = doc.CreateElement("Inner");
                    elemInner.InnerText = exception.InnerException.StackTrace;
                    root.AppendChild(elemInner);
                }

                string nombreArchivo = string.Format("{0:dd_MM_yyyy - HH_mm_ss}", DateTime.Now) + " - " + Guid.NewGuid();
                StreamWriter outStream = File.CreateText(AppDomain.CurrentDomain.BaseDirectory + "\\" + dirLog + "\\" + nombreArchivo + ".xml");
                doc.Save(outStream);

                outStream.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
