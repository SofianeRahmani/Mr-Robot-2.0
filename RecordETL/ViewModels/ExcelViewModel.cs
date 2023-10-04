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
            MembresIndexes = new List<AttributeIndex>();
            Type type = typeof(Membre);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row" || property.Name == "Transactions") continue;

                MembresIndexes.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }


            TransactionsIndexes = new List<AttributeIndex>();
            type = typeof(Transaction);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row") continue;

                TransactionsIndexes.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
            }


            EmployeursIndexes = new List<AttributeIndex>();
            type = typeof(Employeur);
            foreach (var property in type.GetProperties())
            {
                if (property.Name == "Row") continue;

                EmployeursIndexes.Add(new AttributeIndex() { Name = property.Name, Index = -1 });
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




        private List<AttributeIndex> _membresIndexes = new List<AttributeIndex>();
        public List<AttributeIndex> MembresIndexes
        {
            get => _membresIndexes;
            set => SetField(ref _membresIndexes, value);
        }


        private List<string> _membresColumns = new List<string>();
        public List<string> MembresColumns
        {
            get => _membresColumns;
            set => SetField(ref _membresColumns, value);
        }


        private MembresSet _membresSet = new MembresSet();
        public MembresSet MembresSet
        {
            get => _membresSet;
            set => SetField(ref _membresSet, value);
        }




        private List<AttributeIndex> _employeursIndexes = new List<AttributeIndex>();
        public List<AttributeIndex> EmployeursIndexes
        {
            get => _employeursIndexes;
            set => SetField(ref _employeursIndexes, value);
        }


        private List<string> _employeurColumns = new List<string>();
        public List<string> EmployeurColumns
        {
            get => _employeurColumns;
            set => SetField(ref _employeurColumns, value);
        }


        private EmployeursSet _employeursSet = new EmployeursSet();
        public EmployeursSet EmployeursSet
        {
            get => _employeursSet;
            set => SetField(ref _employeursSet, value);
        }

        public ICommand? ExtractCommand
        {
            get
            {
                return
                    new RelayCommand(execute: _ =>
                    {
                        MembresColumns = new List<string>();
                        MembresSet = new MembresSet();
                        
                        TransactionsColumns = new List<string>();
                        TransactionsSet = new TransactionsSet();

                        EmployeurColumns = new List<string>();
                        EmployeursSet = new EmployeursSet();

                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        var fileInfo = new FileInfo(ExcelPath);
                        using var package = new ExcelPackage(fileInfo);

                        var workbook = package.Workbook;

                        MembresColumns = MembresService.ReadColumnsNames(workbook);
                        TransactionsColumns = TransactionsService.ReadColumnsNames(workbook);
                        EmployeurColumns = EmployeursService.ReadColumnsNames(workbook);
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

                    var recordSet = MembresService.ReadAndValidate(Workbook, MembresIndexes, IsAmerican, TerminaisonCourriel);
                    MembresSet = MembresService.Validate(recordSet);

                    TransactionsSet = TransactionsService.ReadAndValidate(Workbook, TransactionsIndexes);


                    EmployeursSet = EmployeursService.ReadAndValidate(Workbook, EmployeursIndexes);

                });
            }
        }
    }
}
