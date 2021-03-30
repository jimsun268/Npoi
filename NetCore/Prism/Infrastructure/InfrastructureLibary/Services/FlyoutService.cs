using InfrastructureLibary.Commands;
using InfrastructureLibary.Constants;
using InfrastructureLibary.IServices;
using MahApps.Metro.Controls;
using Prism.Commands;
using Prism.Regions;
using System.Linq;
using System.Windows.Input;

namespace InfrastructureLibary.Services
{
    public class FlyoutService : IFlyoutService
    {
        readonly IRegionManager _regionManager;
        readonly IApplicationCommands _applicationCommands;
        public ICommand ShowFlyoutCommand { get; private set; }
        public ICommand NavigateRegionCommand { get; private set; }
        public FlyoutService(IRegionManager regionManager, IApplicationCommands applicationCommands)
        {
            _regionManager = regionManager;
            _applicationCommands = applicationCommands;

            this.ShowFlyoutCommand = new DelegateCommand<string>(ShowFlyout);
            this.NavigateRegionCommand = new DelegateCommand<string>(NavigateRegion);
            //注册子命令给全局复合命令
            _applicationCommands.ShowCommand.RegisterCommand(this.ShowFlyoutCommand);
            _applicationCommands.NavigateCommand.RegisterCommand(this.NavigateRegionCommand);

        }
        public void NavigateRegion(string regionName)
        {
            IRegion region = _regionManager.Regions[RegionNames.MainShowRegion];
            if (region != null)
            {
                region.RequestNavigate(regionName);
            }
        }
        public void ShowFlyout(string flyoutName)
        {
            var region = _regionManager.Regions[RegionNames.FlyoutRegion];

            if (region != null)
            {
                var flyout = region.Views.Where(v => v is IFlyoutView && ((IFlyoutView)v).FlyoutName.Equals(flyoutName)).FirstOrDefault() as Flyout;

                if (flyout != null)
                {
                    flyout.IsOpen = !flyout.IsOpen;
                }
            }
        }
    }
}
