using InfrastructureLibary.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureLibary.DataBase
{
    public class efDbContext : DbContext
    {
        //public DbSet<Contract> Contract { get; set; }
        //public DbSet<Land> Land { get; set; }

        //public DbSet<Personnel> Personnel { get; set; }

        public DbSet<LandInformation> LandInformation { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

        }
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            options.UseSqlite("Data Source=blogging.db");
            
        }
    }
}

//add-migration initialcreate
//update-database
