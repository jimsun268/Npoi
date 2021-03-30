using InfrastructureLibary.Commands;
using InfrastructureLibary.Constants;
using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp.ViewModels.FlyoutsRegion
{
    public class EditToolFloutViewModel : BindableBase
    {
        public EditToolFloutViewModel(IApplicationCommands applicationCommands, IEditCommands editCommands, IRegionManager regionManager)
        {
            this.ApplicationCommands = applicationCommands;
            _regionManager = regionManager;
            EditCommands = editCommands;

        }

        #region Fields

        private readonly IRegionManager _regionManager;
        private IApplicationCommands _applicationCommands;
        IRegion _navigationRegion;
        private bool _isCanExcute;
        private IEditCommands _editCommands;
        #endregion

        #region Properties
        public IRegion NavigationRegion => _navigationRegion ??=
                                           _regionManager.Regions[RegionNames.MainShowRegion];
        public bool IsCanExcute
        {
            get { return _isCanExcute; }
            set { SetProperty(ref _isCanExcute, value); }
        }

        #endregion

        #region Commands

        public IEditCommands EditCommands
        {
            get { return _editCommands; }
            set { SetProperty(ref _editCommands, value); }
        }
        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set { SetProperty(ref _applicationCommands, value); }
        }

        private DelegateCommand _goForwardCommand;
        public DelegateCommand GoForwardCommand =>
            _goForwardCommand ?? (_goForwardCommand = new DelegateCommand(ExecuteGoForwardCommand));

        private DelegateCommand _goBackCommand;
        public DelegateCommand GoBackCommand =>
            _goBackCommand ?? (_goBackCommand = new DelegateCommand(ExecuteGoBackCommand));

        #endregion

        #region Excutes

        private void ExecuteGoBackCommand()
        {
            NavigationRegion.NavigationService.Journal.GoBack();
        }
        private void ExecuteGoForwardCommand()
        {
            NavigationRegion.NavigationService.Journal.GoForward();
        }
        #endregion
    }
}
