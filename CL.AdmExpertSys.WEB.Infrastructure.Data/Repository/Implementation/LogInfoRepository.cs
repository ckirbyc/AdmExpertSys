
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;
using System.Data.Entity;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class LogInfoRepository : Repository<LOG_INFO>, ILogInfoRepository
    {
        public LogInfoRepository(DbContext context)
            : base(context)
        {
        }
    }
}
