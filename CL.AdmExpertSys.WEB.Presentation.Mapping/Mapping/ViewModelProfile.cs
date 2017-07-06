
using AutoMapper;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using CL.AdmExpertSys.WEB.Presentation.ViewModel;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Mapping
{
    public class ViewModelProfile : Profile
    {
        protected override void Configure()
        {
            Mapper.CreateMap<AccionInfo, AccionInfoVm>().ReverseMap();
            Mapper.CreateMap<LogInfo, LogInfoVm>().ReverseMap();
        }

        public override string ProfileName
        {
            get { return GetType().Name; }
        }
    }
}
