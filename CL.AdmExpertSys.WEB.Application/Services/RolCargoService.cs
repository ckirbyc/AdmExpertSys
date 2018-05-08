using CL.AdmExpertSys.WEB.Application.Contracts.Services;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;

namespace CL.AdmExpertSys.WEB.Application.Services
{
    public class RolCargoService : BaseService<ROL_CARGO>, IRolCargoService
    {
        protected IRolCargoRepository Repo;

        public RolCargoService(IRolCargoRepository repo)
            : base(repo)
        {
            Repo = repo;
        }
    }
}
