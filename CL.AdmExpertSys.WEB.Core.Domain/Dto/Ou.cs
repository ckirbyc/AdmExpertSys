
using System;

namespace CL.AdmExpertSys.WEB.Core.Domain.Dto
{
    [Serializable]
    public class Ou
    {
        public Ou()
        {
            IdOu = 0;
            IdPadreOu = 0;
        }

        public int IdOu { get; set; }
        public int IdPadreOu { get; set; }
        public string Nombre { get; set; }
        public int Nivel { get; set; }
        public string Ldap { get; set; }
    }
}
