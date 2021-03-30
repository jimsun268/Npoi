using Prism.Commands;
using System;

namespace InfrastructureLibary.Commands
{
    public class ApplicationCommands : IApplicationCommands
    {
        private CompositeCommand _showCommand = new CompositeCommand();
        public CompositeCommand ShowCommand
        {
            get { return _showCommand; }
        }

        public CompositeCommand _navigateCommand = new CompositeCommand();

        public CompositeCommand NavigateCommand
        {
            get { return _navigateCommand; }
        }
    }
}
