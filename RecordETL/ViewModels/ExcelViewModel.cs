namespace RecordETL.ViewModels
{
    public class ExcelViewModel : ViewModelBase
    {
        private string? _excelPath = "Drag and Drop Excel File Here";

        public string? ExcelPath
        {
            get => _excelPath; 
            set => SetField(ref _excelPath, value);
        }

    }
}
