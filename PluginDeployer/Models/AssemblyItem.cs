using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PluginDeployer.Models
{
    class AssemblyItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        private string _displayName;
        public string DisplayName
        {
            get { return _displayName; }
            set
            {
                if (_displayName == value) return;

                _displayName = value;
                OnPropertyChanged();
            }
        }
        public Guid AssemblyId { get; set; }
        public string BoundProject { get; set; }
        public Version Version { get; set; }
        public bool IsWorkflowActivity { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
