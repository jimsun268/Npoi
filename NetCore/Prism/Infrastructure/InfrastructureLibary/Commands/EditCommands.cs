using Prism.Commands;

namespace InfrastructureLibary.Commands
{
    public class EditCommands : IEditCommands
    {
        private CompositeCommand _showCommand = new CompositeCommand();
        public CompositeCommand ShowCommand
        {
            get { return _showCommand; }
        }
        private CompositeCommand _navigateCommand = new CompositeCommand();
        public CompositeCommand NavigateCommand
        {
            get { return _navigateCommand; }
        }
    }
}
