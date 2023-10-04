using RecordETL.Models;
using RecordETL.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Input;

namespace RecordETL.ViewModels
{
    public class ExcelViewModel : ViewModelBase
    {

        public ExcelViewModel()
        {
            SelectedColumns = new List<AttributeIndex>();
            Type type = typeof(Membre);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row" || property.Name == "Transactions") continue;

                SelectedColumns.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }
        }

        
        private bool _isAmerican = false;
        public bool IsAmerican
        {
            get => _isAmerican;
            set => SetField(ref _isAmerican, value);
        }

        private string _terminaisonCourriel;
        public string TerminaisonCourriel
        {
            get => _terminaisonCourriel;
            set => SetField(ref _terminaisonCourriel, value);
        }

        private string? _excelPath = @"";
        public string? ExcelPath
        {
            get => _excelPath;
            set => SetField(ref _excelPath, value);
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


        private MembresSet _membresSet = new MembresSet();
        public MembresSet MembresSet
        {
            get => _membresSet;
            set => SetField(ref _membresSet, value);
        }

        public ICommand? ExtractCommand
        {
            get
            {
                return
                    new RelayCommand(execute: _ =>
                    {
                        AvailableColumns = new List<string>();
                        MembresSet = new MembresSet();
                        AvailableColumns = ExtractorService.ReadDataSourceColumnsNames(ExcelPath);
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
                    var recordSet = ExtractorService.ExtractMembres(ExcelPath, SelectedColumns, IsAmerican, TerminaisonCourriel);
                    MembresSet = ValidatorService.Validate(recordSet);

                    Debug.WriteLine("Validation done");
                });
            }
        }
    }
}
