using CommonServiceLocator;
using InfrastructureLibary.Commands;
using InfrastructureLibary.CustomerRegionAdapters;
using InfrastructureLibary.IServices;
using InfrastructureLibary.Services;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using Prism.Unity;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;
using WpfApp.HelpClass;
using WpfApp.ViewModels.Dialogs;
using WpfApp.Views.Dialogs;

namespace WpfApp
{
    public partial class App : PrismApplication
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
        }
        protected override Window CreateShell()
        {
            return Container.Resolve<Views.MainWindow>();
        }
        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            //注册服务
            //containerRegistry.Register<IPersonnelService, PersonnelService>();
            //containerRegistry.Register<IContractService, ContractService>();


            //注册全局命令
            containerRegistry.RegisterSingleton<IApplicationCommands, ApplicationCommands>();
            containerRegistry.RegisterSingleton<IEditCommands, EditCommands>();
            containerRegistry.RegisterInstance<IFlyoutService>(Container.Resolve<FlyoutService>());

            //注册导航
            //containerRegistry.RegisterForNavigation<MainShow>();


            //注册对话框
            containerRegistry.RegisterDialog<AlertDialog, AlertDialogViewModel>();
            containerRegistry.RegisterDialog<SuccessDialog, SuccessDialogViewModel>();
            containerRegistry.RegisterDialog<WarningDialog, WarningDialogViewModel>();
            containerRegistry.RegisterDialogWindow<DialogWindow>();


            ServiceLocator.SetLocatorProvider(() => new UnityServiceLocator(this.Container.GetContainer()));
        }
        protected override void ConfigureRegionAdapterMappings(RegionAdapterMappings regionAdapterMappings)
        {
            base.ConfigureRegionAdapterMappings(regionAdapterMappings);
            regionAdapterMappings.RegisterMapping(typeof(UniformGrid), Container.Resolve<UniformGridRegionAdapter>());
        }
        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            //moduleCatalog.AddModule<LandInformationModule.LandInformationModule>();
        }
        protected override IModuleCatalog CreateModuleCatalog()
        {
            return new DirectoryModuleCatalog() { ModulePath = @".\" };
        }
    }
}
