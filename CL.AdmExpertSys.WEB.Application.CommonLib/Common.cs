
using System.Configuration;
using System.Globalization;

namespace CL.AdmExpertSys.WEB.Application.CommonLib
{
    public class Common
    {
        #region Métodos de Validación de CL_RUN


        /// <summary>
        /// Valida DV de un CL_RUN
        /// </summary>
        /// <param name="iNumber">Número del CL_RUN</param>
        /// <param name="sDv">Dígito verificador CL_RUN</param>
        /// <returns>
        /// Retorna True si el CL_RUN es válido
        /// </returns>
        public bool ValidateCL_RUN(int iNumber, string sDv)
        {
            if (sDv == CalculateCL_RUN_DV(iNumber))
            {
                return true;
            }
            return false;
        }


        /// <summary>
        /// Calcula dígito verficador (DV) de un CL_RUN
        /// </summary>
        /// <param name="iNumber">Número del CL_RUN</param>
        /// <returns>
        /// Retorna Dígito Verficador del número dado
        /// </returns>
        private static string CalculateCL_RUN_DV(int iNumber)
        {
            string sNumber = iNumber.ToString(CultureInfo.InvariantCulture);
            int calc = 0;

            int[] fact = { 3, 2, 7, 6, 5, 4, 3, 2 };
            int iFact = fact.Length - 1;

            for (int i = sNumber.Length - 1; i >= 0; i--)
            {
                calc += (fact[iFact] * int.Parse(sNumber.Substring(i, 1)));
                iFact--;
            }

            string sDv;
            int resultado = 11 - (calc % 11);

            if (resultado == 11)
            {
                sDv = "0";
            }
            else if (resultado == 10)
            {
                sDv = "K";
            }
            else
            {
                sDv = resultado.ToString(CultureInfo.InvariantCulture);
            }

            return sDv;
        }

        #endregion

        #region Métodos Privados Archivos de Configuración


        /// <summary>
        /// Obtiene un valor de la sección AppSettings del archivo de configuración dada una
        /// llave
        /// </summary>
        /// <param name="pKey">Llave de bpusqueda en el appSettings</param>
        /// <returns>
        /// Retorna el valor de una llave de la sección AppSettings del archivo de
        /// configuración
        /// </returns>
        public string GetAppSetting(string pKey)
        {
            string sRet;
            try
            {
                sRet = ConfigurationManager.AppSettings.Get(pKey);
            }
            catch
            {
                sRet = string.Empty;
            }
            return sRet;
        }
        #endregion

        #region Métodos varios - herramientas
        /// <summary>
        ///  Formatea un string dejando la primera letra de cada palabra en mayusculas
        /// </summary>
        /// <param name="value">String de entrada</param>
        /// <returns>
        /// Retorna el string con cada palabra con la primera letra mayúscula
        /// </returns>
        public string UppercaseWords(string value)
        {
            char[] array = value.ToCharArray();
            // Handle the first letter in the string.
            if (array.Length >= 1)
            {
                if (char.IsLower(array[0]))
                {
                    array[0] = char.ToUpper(array[0]);
                }
            }
            // Scan through the letters, checking for spaces.
            // ... Uppercase the lowercase letters following spaces.
            for (int i = 1; i < array.Length; i++)
            {
                if (array[i - 1] == ' ')
                {
                    if (char.IsLower(array[i]))
                    {
                        array[i] = char.ToUpper(array[i]);
                    }
                }
            }
            return new string(array);
        }

        #endregion
    }
}
