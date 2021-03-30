using InfrastructureLibary.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp.ViewModels.FlyoutsRegion
{
    public class MainFloutViewModel : BindableBase
    {
        public MainFloutViewModel(IApplicationCommands applicationCommands, IRegionManager regionManager)
        {
            this.ApplicationCommands = applicationCommands;
            _regionManager = regionManager;
        }

        #region Fields

        private readonly IRegionManager _regionManager;
        private IApplicationCommands _applicationCommands;


        #endregion

        #region Properties




        #endregion

        #region Commands

        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set { SetProperty(ref _applicationCommands, value); }
        }

        #endregion

        #region Excutes


        #endregion
    }
}
