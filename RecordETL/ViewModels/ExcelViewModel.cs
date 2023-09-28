using RecordETL.Models;
using RecordETL.Services;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Input;

namespace RecordETL.ViewModels
{
    public class ExcelViewModel : ViewModelBase
    {
        public ExcelViewModel()
        {
            SelectedColumns = new List<AttributeIndex>
            {
                new AttributeIndex() { Name = "NumeroMembre", Index = -1 },
                new AttributeIndex() { Name = "Nom", Index = -1 },
                new AttributeIndex() { Name = "Prenom", Index = -1 },
                new AttributeIndex() { Name = "Sexe", Index = -1 },
                new AttributeIndex() { Name = "CourrielTravail", Index = -1 },
                new AttributeIndex() { Name = "CourrielPersonnel", Index = -1 },
                new AttributeIndex() { Name = "CourrielAutre", Index = -1 },
                new AttributeIndex() { Name = "Telephone", Index = -1 },
                new AttributeIndex() { Name = "TelephoneTravail", Index = -1 },
                new AttributeIndex() { Name = "TelephoneCellulaire", Index = -1 },
                new AttributeIndex() { Name = "Adresse", Index = -1 },
                new AttributeIndex() { Name = "Ville", Index = -1 },
                new AttributeIndex() { Name = "Province", Index = -1 },
                new AttributeIndex() { Name = "CodePostal", Index = -1 },
                new AttributeIndex() { Name = "Nas", Index = -1 },
                new AttributeIndex() { Name = "Categories", Index = -1 },
                new AttributeIndex() { Name = "DateNaissance", Index = -1 },
                new AttributeIndex() { Name = "DateAnciennete", Index = -1 },
                new AttributeIndex() { Name = "Anciennete", Index = -1 },
                new AttributeIndex() { Name = "DateEmbauche", Index = -1 },
                new AttributeIndex() { Name = "Statut", Index = -1 },
                new AttributeIndex() { Name = "DateStatut", Index = -1 },
                new AttributeIndex() { Name = "IdSystemeSource", Index = -1 },
                new AttributeIndex() { Name = "Secteur", Index = -1 },
                new AttributeIndex() { Name = "StatutPersonne", Index = -1 },
                new AttributeIndex() { Name = "IdentifiantAlternatif", Index = -1 },
                new AttributeIndex() { Name = "InfosComplementaires1", Index = -1 },
                new AttributeIndex() { Name = "InfosComplementaires2", Index = -1 },                
                
                new AttributeIndex() { Name = "Employeur", Index = -1 },
                new AttributeIndex() { Name = "NumeroEmployeur", Index = -1 },
                new AttributeIndex() { Name = "Fonction", Index = -1 },
                new AttributeIndex() { Name = "DateDebut", Index = -1 },
                new AttributeIndex() { Name = "DateFin", Index = -1 },
                new AttributeIndex() { Name = "InfosComplementairesEmplois", Index = -1 }
            };

        }

        private string? _excelPath = @"C:\Users\Bucket\Desktop\sample.xlsx";

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
                    var recordSet = ExtractorService.Extract(ExcelPath, SheetIndex, SelectedColumns);
                    RecordSet = ValidatorService.Validate(recordSet);
                });
            }
        }
    }
}
