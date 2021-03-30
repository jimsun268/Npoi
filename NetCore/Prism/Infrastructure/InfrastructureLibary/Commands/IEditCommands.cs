using Prism.Commands;

namespace InfrastructureLibary.Commands
{
    public interface IEditCommands
    {
        public CompositeCommand ShowCommand { get; }
        public CompositeCommand NavigateCommand { get; }

    }
}
