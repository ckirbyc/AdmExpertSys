
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class AccionInfoService : BaseService<AccionInfo>, IAccionInfoService
    {
        protected IAccionInfoRepository Repo;

        public AccionInfoService(IAccionInfoRepository repo)
            :base(repo)
        {
            Repo = repo;
        }
    }
}
