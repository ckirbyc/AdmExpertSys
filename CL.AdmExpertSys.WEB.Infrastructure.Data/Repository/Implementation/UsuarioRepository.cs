using CL.AdmExpertSys.WEB.Core.Domain.Contracts.Repository;
using CL.AdmExpertSys.WEB.Core.Domain.Model;
using Pragma.Commons.Data;
using System.Data.Entity;

namespace CL.AdmExpertSys.WEB.Infrastructure.Data.Repository.Implementation
{
    public class UsuarioRepository : Repository<USUARIO>, IUsuarioRepository
    {
        public UsuarioRepository(DbContext context)
            : base(context)
        {
        }
    }
}
