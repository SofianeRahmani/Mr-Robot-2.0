using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace RecordETL.Services
{
    public class ExtractorService
    {

        public static List<string> ReadColumnsNames(string filePath, int pageIndex)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<string> columns = new List<string>();

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[pageIndex];


            if (worksheet.Dimension == null) return columns; // Return empty list if worksheet is empty


            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                columns.Add(worksheet.Cells[3, col].Text);
            }

            return columns;
        }

        public static string? GetColumnValue(int row, int column, ExcelWorksheet worksheet)
        {
            return column != -1 ? worksheet.Cells[row, column + 1].Text : null;
        }

        public static RecordSet Extract(string filePath, int pageIndex, List<AttributeIndex> positions, bool isAmerican)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            RecordSet recordSet = new RecordSet();
            recordSet.Records = new List<Record>();
            recordSet.Errors = new List<Models.Error>();

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[pageIndex];


            for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
            {
                var person = new Record();
                Type type = typeof(Record);
                person.Row = row;

                foreach (var position in positions)
                {
                    PropertyInfo propertyInfo = type.GetProperty(position.Name);
                    propertyInfo.SetValue(person, GetColumnValue(row, position.Index, worksheet));
                }

                person.Nom = person.Nom?.Trim() + " " + person.SecondNom;
                person.Prenom = person.Prenom?.Trim() + " " + person.SecondPrenom;

                person.Telephone = FormatPhoneNumber(person.Telephone);
                person.TelephoneTravail = FormatPhoneNumber(person.TelephoneTravail);
                person.TelephoneCellulaire = FormatPhoneNumber(person.TelephoneCellulaire);
                person.CodePostal = FormatPostalCode(person.CodePostal, isAmerican);



                recordSet.Records.Add(person);
            }



            return recordSet;
        }


        static string FormatPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
            {
                return phoneNumber;
            }

            phoneNumber = Regex.Replace(phoneNumber, @"\D", "");
            if (phoneNumber.Any(c => !char.IsDigit(c)))
            {
                return phoneNumber;
            }

            return $"{phoneNumber.Substring(0, 3)}-{phoneNumber.Substring(3, 3)}-{phoneNumber.Substring(6, 4)}";
        }


        static string FormatPostalCode(string postalCode, bool isAmerican)
        {
            if (string.IsNullOrEmpty(postalCode))
            {
                return postalCode;
            }

            // Suppression de tous les caractères non alphanumériques
            postalCode = Regex.Replace(postalCode, @"[^\w]", "");

            if (isAmerican)
            {
                if (postalCode.Length != 5 || !Regex.IsMatch(postalCode, @"^\d{5}$"))
                    throw new ArgumentException("Le code postal américain doit être composé de 5 chiffres.");

                return postalCode;
            }
            else
            {
                if (postalCode.Length != 6 || !Regex.IsMatch(postalCode, @"^[A-Za-z]\d[A-Za-z]\d[A-Za-z]\d$"))
                    throw new ArgumentException("Le code postal canadien doit être composé de 6 caractères alphanumériques alternés.");

                return $"{postalCode.Substring(0, 3)} {postalCode.Substring(3, 3)}";
            }
        }
    }
}
