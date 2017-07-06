
using System.ComponentModel;

namespace CL.AdmExpertSys.WEB.Core.Domain.Enums
{
    public enum EnumAccionInfo
    {
        [Description("Buscar")]
        Buscar = 1,
        [Description("Guardar")]
        Guardar = 2,
        [Description("Sincronizar")]
        Sincronizar = 3,
        [Description("Asignar Licencia")]
        AsignarLicencia = 4,
        [Description("Deshabilitar Usuario")]
        DeshabilitarUsuario = 5,
        [Description("Login")]
        Login = 6,
        [Description("Verificar Existencia Usuario")]
        VerificarExistenciaUsuario = 7,
        [Description("Asignar Grupo")]
        AsignarGrupo = 8,
        [Description("Guardar Grupo")]
        GuardarGrupo = 9
    }
}
