using Prism.Commands;

namespace InfrastructureLibary.Commands
{
    public interface IApplicationCommands
    {
        public CompositeCommand ShowCommand { get; }
        public CompositeCommand NavigateCommand { get; }
    }
}
