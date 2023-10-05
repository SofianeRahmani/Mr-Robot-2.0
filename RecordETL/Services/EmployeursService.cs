using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace RecordETL.Services
{
    public class EmployeursService
    {

        public static List<string> ReadColumnsNames(ExcelWorkbook workbook)
        {
            List<string> columns = new List<string>();
            var worksheet = workbook.Worksheets[2];
            if (worksheet.Dimension == null) return columns; // Return empty list if worksheet is empty
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                columns.Add(worksheet.Cells[1, col].Text);
            }

            return columns;
        }



        public static string? GetColumnValue(int row, int column, ExcelWorksheet worksheet)
        {
            return column != -1 ? worksheet.Cells[row, column + 1].Text.Trim() : null;
        }



        public static EmployeursSet ReadAndValidate(ExcelWorkbook workbook,
            List<AttributeIndex> positions)
        {

            EmployeursSet Set = new EmployeursSet();
            Set.Employeurs = new List<Employeur>();
            Set.Errors = new List<Models.Error>();

            var sheet = workbook.Worksheets[2];


            List<string> employeurs = new List<string>();
            var fonctionsSheet = workbook.Worksheets[6];

            for (int row = 2; row <= fonctionsSheet.Dimension.End.Row; row++)
            {
                string value = GetColumnValue(row, 1, fonctionsSheet);

                if (value == null || value == "") break;

                employeurs.Add(value);
            }

            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                bool empty = true;
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    if (sheet.Cells[row, col].Text != "")
                    {
                        empty = false;
                    }
                }

                if (empty) break;

                var employeur = new Employeur();
                Type type = typeof(Employeur);
                employeur.Row = row;

                foreach (var position in positions)
                {
                    PropertyInfo propertyInfo = type.GetProperty(position.Name);
                    propertyInfo.SetValue(employeur, GetColumnValue(row, position.Index, sheet));
                }


                employeur.Telephone = employeur.Telephone != null ? Regex.Replace(employeur.Telephone, @"[^0-9]", "") : null;

                if (!employeurs.Contains(employeur.Nom))
                {
                    var error = new Models.Error()
                    {
                        Code = "ERR-002",
                        Description_EN = "Employer Name do not exists",
                        Description_FR = "Nom Employeur n'est existe pas déjà",
                        RecordIndex = employeur.Row
                    };

                    Set.Errors.Add(error);
                }
                
                Set.Employeurs.Add(employeur);
            }



            var missingNumeroMembre = Set.Employeurs.Where(r => string.IsNullOrEmpty(r.Numero)).ToList();
            for (int i = 0; i < missingNumeroMembre.Count; i++)
            {
                var record = missingNumeroMembre[i];

                int number = i + 1;
                string value = number < 10 ? $"00{number}" : number < 100 ? $"0{number}" : number.ToString();
                record.Numero = $"SN-{value}";


                var error = new Models.Error()
                {
                    Code = "ERR-001",
                    Description_EN = "Employer Number is required",
                    Description_FR = "Numero Employeur est requis",
                    RecordIndex = record.Row
                };

                Set.Errors.Add(error);
            }



            // remove duplicates
            var groupedRecords = from r in Set.Employeurs
                group r by r.Numero into g
                where g.Count() > 1
                select g;

            foreach (var group in groupedRecords)
            {
                var records = group.ToList();
                for (int i = 0; i < records.Count; i++)
                {
                    var record = records[i];
                    if (i > 0) // Do not modify the first membre
                    {
                        record.Numero = $"{record.Numero}-D{i}"; // Append D1, D2, D3, etc. to duplicates
                    }
                }
            }

            return Set;
        }

        internal static void ExportCSV(EmployeursSet employeursSet, string outputPath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{outputPath}\\employeurs.csv"))
            {
                file.WriteLine("Numero,Nom,Adresse,Ville,Province,CodePostal,InformationComplémentaire,Telephone");
                foreach (var record in employeursSet.Employeurs)
                {
                    file.WriteLine($"{record.Numero},{record.Nom},{record.Adresse},{record.Ville},{record.Province},{record.CodePostal},{record.InformationComplémentaire},{record.Telephone}");
                }
            }
        }
    }
}
