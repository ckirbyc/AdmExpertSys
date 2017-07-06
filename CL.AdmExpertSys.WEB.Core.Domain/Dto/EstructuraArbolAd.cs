
using System;

namespace CL.AdmExpertSys.WEB.Core.Domain.Dto
{
    [Serializable]
    public class EstructuraArbolAd
    {
        public EstructuraArbolAd()
        {
            Cabecera = string.Empty;
        }
        public string Cabecera { get; set; }
        public string CodArbolHtml { get; set; }
    }
}
