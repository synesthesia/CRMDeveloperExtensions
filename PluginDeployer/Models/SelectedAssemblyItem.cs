using System.ComponentModel;

namespace PluginDeployer.Models
{
    static class SelectedAssemblyItem
    {
        private static AssemblyItem _item;

        public static AssemblyItem Item
        {
            get { return _item; }
            set
            {
                if (_item == value) return;

                _item = value;
                OnPropertyChanged("Item");
            }
        }

        public static event PropertyChangedEventHandler PropertyChanged;

        public static void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(null, new PropertyChangedEventArgs(name));
            }
        }
    }
}
