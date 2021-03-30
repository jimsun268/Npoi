using InfrastructureLibary.Models;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureLibary.Events
{
    public class LandInformationEvent : PubSubEvent<LandInformation>
    {
        public void IsChange()
        {

        }
    }
}
