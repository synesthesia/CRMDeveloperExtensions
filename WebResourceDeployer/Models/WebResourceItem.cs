using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace WebResourceDeployer.Models
{
    class WebResourceItem : INotifyPropertyChanged
    {
        private bool _publish;
        public bool Publish
        {
            get { return _publish; }
            set
            {
                if (_publish == value) return;

                _publish = value;
                OnPropertyChanged();
            }
        }
        public Guid WebResourceId { get; set; }
        public int Type { get; set; }
        public string TypeName { get; set; }
        public string Name { get; set; }
        public bool IsManaged { get; set; }
        private ObservableCollection<ComboBoxItem> _projectFiles;
        public ObservableCollection<ComboBoxItem> ProjectFiles
        {
            get { return _projectFiles; }
            set
            {
                if (_projectFiles == value) return;

                _projectFiles = value;
                OnPropertyChanged();
            }
        }
        private ObservableCollection<MenuItem> _projectFolders;
        public ObservableCollection<MenuItem> ProjectFolders
        {
            get { return _projectFolders; }
            set
            {
                if (_projectFolders == value) return;

                _projectFolders = value;
                OnPropertyChanged();
            }
        }
        private bool _allowDiff;
        public bool AllowDiff
        {
            get { return _allowDiff; }
            set
            {
                if (_allowDiff == value) return;

                _allowDiff = value;
                OnPropertyChanged();
            }
        }
        private bool _allowPublish;
        public bool AllowPublish
        {
            get { return _allowPublish; }
            set
            {
                if (_allowPublish == value) return;

                _allowPublish = value;
                OnPropertyChanged();
            }
        }
        private string _boundFile;
        public string BoundFile
        {
            get { return _boundFile; }
            set
            {
                if (_boundFile == value) return;

                _boundFile = value;
                OnPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
