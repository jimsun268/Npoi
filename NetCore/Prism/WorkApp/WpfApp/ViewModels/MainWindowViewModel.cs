using InfrastructureLibary.Constants;
using Prism.Commands;
using Prism.Modularity;
using Prism.Mvvm;
using Prism.Regions;
using Prism.Services.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp.Views.FlyoutsRegion;

namespace WpfApp.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        public MainWindowViewModel(IModuleManager moduleManager, IRegionManager regionManager, IDialogService dialogService)
        {
            _moduleManager = moduleManager;
            _regionManager = regionManager;
            _dialogService = dialogService;
            _moduleManager.LoadModuleCompleted += ModuleManager_LoadModuleCompleted;
        }

        #region Fields
        private readonly IModuleManager _moduleManager;
        private readonly IDialogService _dialogService;
        private readonly IRegionManager _regionManager;
        #endregion

        #region Properties


        #endregion

        #region Commands

        private DelegateCommand _loadingCommand;
        public DelegateCommand LoadingCommand =>
               _loadingCommand ??= new DelegateCommand(ExecuteLoadingCommand);
        #endregion

        #region Excutes
        private void ModuleManager_LoadModuleCompleted(object sender, LoadModuleCompletedEventArgs e)
        {

        }
        private void ExecuteLoadingCommand()
        {

            _regionManager.RegisterViewWithRegion(RegionNames.LeftWindowRegion, typeof(LeftWindowRegion));
            _regionManager.RegisterViewWithRegion(RegionNames.FlyoutRegion, typeof(MainFlout));
            _regionManager.RegisterViewWithRegion(RegionNames.FlyoutRegion, typeof(EditToolFlout));

            IRegion region = _regionManager.Regions[RegionNames.MainShowRegion];
            region.RequestNavigate("MainShow", NavigationCompelted);

        }
        private void NavigationCompelted(NavigationResult obj)
        {


        }
        #endregion
    }
}
