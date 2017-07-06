
using System.Data.Entity;
using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class LogInfoRepository : Repository<LogInfo>, ILogInfoRepository
    {
        public LogInfoRepository(DbContext context)
            : base(context)
        {
        }
    }
}
