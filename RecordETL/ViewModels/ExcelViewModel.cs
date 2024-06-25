using OfficeOpenXml;
using RecordETL.Models;
using RecordETL.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;

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
        private string _selectedSheet;
        private ObservableCollection<string> _columnNames = new ObservableCollection<string>();
        public ObservableCollection<string> ColumnNames
        {
            get => _columnNames;
            set => SetField(ref _columnNames, value);
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
            EmploisIndex = new List<AttributeIndex>();
            //type = typeof(emplois);

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
        private List<AttributeIndex> _emploisIndexes = new List<AttributeIndex>();
        
        public List<AttributeIndex> EmploisIndex
        {
            get => _emploisIndexes;
            set => SetField(ref _emploisIndexes, value);
        }
        
        
        private List<string> _emploisColumns = new List<string>();
        public List<string> EmploisColumns
        {
            get => _emploisColumns;
            set => SetField(ref _emploisColumns, value);
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

        public ICommand? ExtractSheetCommand
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
                            List<string> sheetNames = RecordETL.Services.ExcelService.ReadSheetNames(package.Workbook);

                            //MembresColumns = MembresService.ReadColumnsNames(workbook, sheetNames);
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

                    var recordSet =
                        MembresService.ReadAndValidate(Workbook, MembresIndexes, IsAmerican, TerminaisonCourriel);
                    MembresSet = MembresService.Validate(recordSet);
                    TransactionsSet = TransactionsService.ReadAndValidate(Workbook, TransactionsIndexes);
                    EmployeursSet = EmployeursService.ReadAndValidate(Workbook, EmployeursIndexes);
                });
            }
        }

        private string _outputPath = @"";

        public string OutputPath
        {
            get => _outputPath;
            set
            {
                SetField(ref _outputPath, value);

                if (_outputPath != "")
                {
                    try
                    {
                        MembresService.ExportCSV(MembresSet, OutputPath);
                        TransactionsService.ExportCSV(TransactionsSet, OutputPath);
                        EmployeursService.ExportCSV(EmployeursSet, OutputPath);


                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        var fileInfo = new FileInfo(OutputPath + "/Errors.xls");

                        using var package = new ExcelPackage(fileInfo);
                        var workbook = package.Workbook;


                        MembresService.ExportErrors(MembresSet, workbook);
                        TransactionsService.ExportErrors(TransactionsSet, workbook);
                        EmployeursService.ExportErrors(EmployeursSet, workbook);


                        package.SaveAs(fileInfo);
                        package.Dispose();


                        fileInfo = new FileInfo(ExcelPath);
                        using var package2 = new ExcelPackage(fileInfo);

                        workbook = package2.Workbook;

                        ConvertExcelToCsv(workbook.Worksheets[5], OutputPath + "/emplois.csv");
                        ConvertExcelToCsv(workbook.Worksheets[7], OutputPath + "/fonctions.csv");
                        ConvertExcelToCsv(workbook.Worksheets[8], OutputPath + "/secteurs.csv");
                        ConvertExcelToCsv(workbook.Worksheets[9], OutputPath + "/pastilles.csv");

                        package.Dispose();

                        // Count the errors
                        int dataCount = MembresSet.Records.Count + TransactionsSet.Transactions.Count +
                                        EmployeursSet.Employeurs.Count;
                        int errorsCount = MembresSet.Errors.Count + TransactionsSet.Errors.Count +
                                          EmployeursSet.Errors.Count;
                        int correctCount = dataCount - errorsCount;
                        int indice = 100 - (errorsCount * 100 / dataCount);

                        var indiceFileInfo = new FileInfo(OutputPath + "/Indice.xls");

                        var package1 = new ExcelPackage(indiceFileInfo);
                        var workbook1 = package1.Workbook;


                        var worksheet = workbook1.Worksheets.Add("Page 1");
                        worksheet.Cells["A1"].Value = "Date et heure";
                        worksheet.Cells["B1"].Value = "Nom du client";
                        worksheet.Cells["C1"].Value = "Nombre total de données";
                        worksheet.Cells["D1"].Value = "Données correcte";
                        worksheet.Cells["E1"].Value = "Données erronées";
                        worksheet.Cells["F1"].Value = "Indice de qualité";


                        worksheet.Cells[$"A2"].Value = DateTime.Now.ToString("dd-MM-yyyy HH:ss");
                        worksheet.Cells[$"B2"].Value = "";
                        worksheet.Cells[$"C2"].Value = dataCount;
                        worksheet.Cells[$"D2"].Value = correctCount;
                        worksheet.Cells[$"E2"].Value = errorsCount;
                        worksheet.Cells[$"F2"].Value = indice;


                        package1.SaveAs(indiceFileInfo);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);

                        MessageBox.Show(e.Message, "Impossible de générer les fichiers", MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
        }


        public void ConvertExcelToCsv(ExcelWorksheet worksheet, string csvFilePath)
        {
            // Create a new CSV file to write to
            using (var sw = new StreamWriter(csvFilePath))
            {
                for (int rowNum = 1; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var line = new List<string>();
                    for (int colNum = 1; colNum <= worksheet.Dimension.End.Column; colNum++)
                    {
                        var cellValue = worksheet.Cells[rowNum, colNum].Text;


                        if (cellValue.Contains(",") || cellValue.Contains("\n"))
                        {
                            cellValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                        }

                        line.Add(cellValue);
                    }

                    var lineStr = string.Join(",", line);
                    sw.WriteLine(lineStr);
                }
            }
        }
    }
}