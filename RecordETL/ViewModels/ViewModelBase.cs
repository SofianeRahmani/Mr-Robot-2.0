using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RecordETL.ViewModels
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        private ObservableCollection<string> _sheetNames = new ObservableCollection<string>();
        public ObservableCollection<string> SheetNames
        {
            get => _sheetNames;
            set
            {
                if (_sheetNames != value)
                {
                    _sheetNames = value;
                    OnPropertyChanged(nameof(SheetNames));
                }
            }
        }
        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
