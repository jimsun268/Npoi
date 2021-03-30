using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureLibary.Models
{
    public class LandInformation
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int Number1 { get; set; }
        public int Number2 { get; set; }
        public int Number3 { get; set; }
        public string Number4 { get; set; }
        public string State { get; set; }
        public int Area { get; set; }
        public int Price { get; set; }
        public string Note { get; set; }
    }

    public class LandInformationConfiguration : IEntityTypeConfiguration<LandInformation>
    {
        public void Configure(EntityTypeBuilder<LandInformation> builder)
        {
            builder.HasKey(t => t.ID);
            builder.HasIndex(t => t.Name);
            builder.Property(t => t.Name).IsRequired();
        }
    }
}
