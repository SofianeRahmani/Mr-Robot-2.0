using RecordETL.Models;
using RecordETL.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Input;
using OfficeOpenXml;

namespace RecordETL.ViewModels
{
    public class ExcelViewModel : ViewModelBase
    {

        private string? _excelPath = @"";
        public string? ExcelPath
        {
            get => _excelPath;
            set => SetField(ref _excelPath, value);
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

        public ExcelViewModel()
        {
            DataSourceIndexes = new List<AttributeIndex>();
            Type type = typeof(Membre);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row" || property.Name == "Transactions") continue;

                DataSourceIndexes.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }


            TransactionsIndexes = new List<AttributeIndex>();
            type = typeof(Transaction);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row") continue;

                TransactionsIndexes.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }
        }

        private List<AttributeIndex> _transactionsIndexes;
        public List<AttributeIndex> TransactionsIndexes
        {
            get => _transactionsIndexes; 
            set => SetField(ref _transactionsIndexes, value);
        }

        private List<string> _transactionsColumns = new List<string>();
        public List<string> TransactionsColumns
        {
            get => _transactionsColumns;
            set => SetField(ref _transactionsColumns, value);
        }


        private TransactionsSet _transactionsSet = new TransactionsSet();
        public TransactionsSet TransactionsSet
        {
            get => _transactionsSet;
            set => SetField(ref _transactionsSet, value);
        }


        private List<AttributeIndex> _dataSourceIndexes = new List<AttributeIndex>();
        public List<AttributeIndex> DataSourceIndexes
        {
            get => _dataSourceIndexes;
            set => SetField(ref _dataSourceIndexes, value);
        }


        private List<string> _dataSourceColumns = new List<string>();
        public List<string> DataSourceColumns
        {
            get => _dataSourceColumns;
            set => SetField(ref _dataSourceColumns, value);
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
                        DataSourceColumns = new List<string>();
                        MembresSet = new MembresSet();

                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        var fileInfo = new FileInfo(ExcelPath);
                        using var package = new ExcelPackage(fileInfo);

                        var Workbook = package.Workbook;

                        DataSourceColumns = DataSourceService.ReadColumnsNames(Workbook);
                        TransactionsColumns = TransactionsService.ReadColumnsNames(Workbook);
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
                    
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var fileInfo = new FileInfo(ExcelPath);
                    using var package = new ExcelPackage(fileInfo);

                    var Workbook = package.Workbook;

                    var recordSet = DataSourceService.ReadAndValidate(Workbook, DataSourceIndexes, IsAmerican, TerminaisonCourriel);
                    MembresSet = DataSourceService.Validate(recordSet);

                    TransactionsSet = TransactionsService.ReadAndValidate(Workbook, TransactionsIndexes);

                });
            }
        }
    }
}
