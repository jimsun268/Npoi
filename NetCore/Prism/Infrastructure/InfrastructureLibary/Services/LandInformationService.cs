using InfrastructureLibary.Constants;
using InfrastructureLibary.IServices;
using InfrastructureLibary.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureLibary.Services
{
    public class LandInformationService : ILandInformationService
    {
        public List<LandInformation> GetAllData()
        {
            return Global.DataBase.LandInformation.AsNoTracking().ToList<LandInformation>();
        }
        public int SaveDataBase()
        {
            return Global.DataBase.SaveChanges();
        }
        public List<EntityEntry<LandInformation>> WriteList(List<LandInformation> landInformations)
        {
            List<EntityEntry<LandInformation>> entities = new List<EntityEntry<LandInformation>>();
            if (landInformations != null)
            {

                foreach (LandInformation li in landInformations)
                {
                    entities.Add(Global.DataBase.LandInformation.Add(li));
                }
            }
            return entities;
        }
        public LandInformation GetDataById(int id)
        {
            return Global.DataBase.LandInformation.Single<LandInformation>(t => t.ID == id);
        }
        public List<LandInformation> GetDataByName(string name)
        {
            return Global.DataBase.LandInformation.Where<LandInformation>(t => t.Name == name).ToList<LandInformation>();
        }
    }
}
