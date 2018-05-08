using CL.AdmExpertSys.Web.Infrastructure.LogTransaccional;
using CL.AdmExpertSys.WEB.Presentation.ViewModel.Resource;
using System;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CL.AdmExpertSys.WEB.Presentation.ViewModel.Validation
{
    public class ValidaCorreosAttribute : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            if (value == null)
                return ValidationResult.Success;

            const string separadorDeCorreos = ";";
            var correos = (string)value;

            var listaDeCorreos = correos.Split(new[] { separadorDeCorreos }, StringSplitOptions.RemoveEmptyEntries);
            var validadorEmail = new ValidadorEmail();

            foreach (var correo in listaDeCorreos)
            {
                var correoValido = validadorEmail.IsValidEmail(correo);

                if (!correoValido)
                    return new ValidationResult(string.Format(AesMensajes.ValidaCorreosAttribute_Invalid, correo, separadorDeCorreos));
            }

            return ValidationResult.Success;
        }
    }

    public class ValidadorEmail
    {
        bool _invalid;

        public bool IsValidEmail(string strIn)
        {
            _invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Use IdnMapping class to convert Unicode domain names. 
            try
            {
                strIn = Regex.Replace(strIn, @"(@)(.+)$", DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));
            }
            catch (RegexMatchTimeoutException ex)
            {
                Utils.LogErrores(ex);
                return false;
            }

            if (_invalid)
                return false;

            // Return true if strIn is in valid e-mail format. 
            try
            {
                return Regex.IsMatch(strIn,
                      @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                      @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                      RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException ex)
            {
                Utils.LogErrores(ex);
                return false;
            }
        }

        private string DomainMapper(Match match)
        {
            // IdnMapping class with default property values.
            var idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            try
            {
                domainName = idn.GetAscii(domainName);
            }
            catch (ArgumentException ex)
            {
                Utils.LogErrores(ex);
                _invalid = true;
            }
            return match.Groups[1].Value + domainName;
        }

    }
}
