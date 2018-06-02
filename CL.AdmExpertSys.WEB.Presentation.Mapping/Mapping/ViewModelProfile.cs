
using AutoMapper;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Mapping
{
    public class ViewModelProfile : Profile
    {
        protected override void Configure()
        {
            Mapper.CreateMap<ACCION_INFO, AccionInfoVm>().ReverseMap();
            Mapper.CreateMap<LOG_INFO, LogInfoVm>().ReverseMap();
            Mapper.CreateMap<PERFIL_USUARIO, PerfilUsuarioVm>().ReverseMap();
            Mapper.CreateMap<ROL_CARGO, RolCargoVm>().ReverseMap();
            Mapper.CreateMap<LICENCIAS_O365, LicenciaO365Vm>().ReverseMap();
            Mapper.CreateMap<MANTENEDOR_LICENCIA, MantenedorLicenciaVm>().ReverseMap();
            Mapper.CreateMap<ESTADO_CUENTA_USUARIO, EstadoCuentaUsuarioVm>().ReverseMap();
            Mapper.CreateMap<USUARIO, UsuarioPerfilVm>().ReverseMap();
        }

        public override string ProfileName
        {
            get { return GetType().Name; }
        }
    }
}
