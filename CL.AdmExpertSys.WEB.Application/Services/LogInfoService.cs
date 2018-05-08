
using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class LogInfoService : BaseService<LOG_INFO>, ILogInfoService
    {
        protected ILogInfoRepository Repo;

        public LogInfoService(ILogInfoRepository repo)
            :base(repo)
        {
            Repo = repo;
        }
    }
}
