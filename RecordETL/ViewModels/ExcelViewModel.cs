using RecordETL.Models;
using RecordETL.Services;
using System;
using System.Collections.Generic;
using System.Windows.Input;

namespace RecordETL.ViewModels
{
    public class ExcelViewModel : ViewModelBase
    {
        public ExcelViewModel()
        {
            SelectedColumns = new List<AttributeIndex>();
            Type type = typeof(Record);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row") continue;

                SelectedColumns.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }

            //AvailableColumns = ExtractorService.ReadColumnsNames(ExcelPath, SheetIndex);
        }

        private bool _isAmerican = false;
        public bool IsAmerican
        {
            get => _isAmerican;
            set => SetField(ref _isAmerican, value);
        }

        private string? _excelPath = @"";
        public string? ExcelPath
        {
            get => _excelPath;
            set => SetField(ref _excelPath, value);
        }

        private int _sheetIndex = 1;
        public int SheetIndex
        {
            get => _sheetIndex;
            set => SetField(ref _sheetIndex, value);
        }

        private List<AttributeIndex> _selectedColumns = new List<AttributeIndex>();
        public List<AttributeIndex> SelectedColumns
        {
            get => _selectedColumns;
            set => SetField(ref _selectedColumns, value);
        }


        private List<string> _availableColumns = new List<string>();
        public List<string> AvailableColumns
        {
            get => _availableColumns;
            set => SetField(ref _availableColumns, value);
        }


        private RecordSet _recordSet = new RecordSet();
        public RecordSet RecordSet
        {
            get => _recordSet;
            set => SetField(ref _recordSet, value);
        }

        public ICommand? ExtractCommand
        {
            get
            {
                return
                    new RelayCommand(execute: _ =>
                    {
                        AvailableColumns = ExtractorService.ReadColumnsNames(ExcelPath, SheetIndex);
                    },
                    o => _excelPath != "");
            }
        }


        public ICommand? ValidateCommand
        {
            get
            {
                return new RelayCommand(execute: _ =>
                {
                    var recordSet = ExtractorService.Extract(ExcelPath, SheetIndex, SelectedColumns, IsAmerican);
                    RecordSet = ValidatorService.Validate(recordSet);
                });
            }
        }
    }
}
