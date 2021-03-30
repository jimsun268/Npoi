using InfrastructureLibary.Models;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureLibary.IServices
{
    public interface ILandInformationService
    {
        List<LandInformation> GetAllData();
        LandInformation GetDataById(int id);
        List<LandInformation> GetDataByName(string name);
        List<EntityEntry<LandInformation>> WriteList(List<LandInformation> landInformations);
        int SaveDataBase();

    }
}
