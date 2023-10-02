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
                columns.Add(worksheet.Cells[1, col].Text);
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


            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                bool empty = true;
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    if (worksheet.Cells[row, col].Text != "")
                    {
                        empty = false;
                    }
                }

                if(empty) break;

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

                if (person.TerminaisonCourriel == null)
                {
                    person.CourrielTravail = person.CourrielTravail?.Trim() + person.TerminaisonCourriel;
                    person.CourrielPersonnel = person.CourrielPersonnel?.Trim() + person.TerminaisonCourriel;
                    person.CourrielAutre = person.CourrielAutre?.Trim() + person.TerminaisonCourriel;
                }
                else
                {
                    if (!person.CourrielTravail.EndsWith(person.TerminaisonCourriel))
                    {
                        person.CourrielAutre = person.CourrielTravail;
                        person.CourrielTravail = null;


                        Error error = new Error()
                        {
                            Code = "ERR-004",
                            Description_EN = "Email address does not match the domain name",
                            Description_FR = "L'adresse e-mail ne correspond pas au nom de domaine",
                            RecordIndex = row
                        };

                        recordSet.Errors.Add(error);
                    }
                }

                if (person.NumeroAppartement != null)
                {
                    person.Adresse = person.NumeroAppartement.Replace("#", "") + " " + person.Adresse?.Trim();
                }




                if (person.CodePostal != null)
                {
                    if (isAmerican)
                    {
                        if (person.CodePostal.Length != 5 || !Regex.IsMatch(person.CodePostal, @"^\d{5}$"))
                        {
                            Error error = new Error()
                            {
                                Code = "ERR-005",
                                Description_EN = "The American postal code must be composed of 5 digits.",
                                Description_FR = "Le code postal américain doit être composé de 5 chiffres.",
                                RecordIndex = row
                            };

                            recordSet.Errors.Add(error);
                        }

                    }
                    else
                    {
                        if (person.CodePostal.Length != 6 || !Regex.IsMatch(person.CodePostal, @"^[A-Za-z]\d[A-Za-z]\d[A-Za-z]\d$"))
                        {
                            Error error = new Error()
                            {
                                Code = "ERR-005",
                                Description_EN = "The Canadian postal code must be composed of 6 characters.",
                                Description_FR = "Le code postal canadien doit être composé de 6 caractères.",
                                RecordIndex = row
                            };

                            recordSet.Errors.Add(error);
                        }
                        else
                        {
                            person.CodePostal = $"{person.CodePostal.Substring(0, 3)} {person.CodePostal.Substring(3, 3)}";
                        }



                    }
                }

                Regex dateRegex = new Regex(@"^\d{4}-\d{2}-\d{2}$");
                if (person.DateNaissance != null)
                {


                    if (!dateRegex.IsMatch(person.DateNaissance))
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-006",
                            Description_EN = "The date of birth format is invalid",
                            Description_FR = "Le format de la date de naissance est invalide",
                            RecordIndex = row
                        };

                        recordSet.Errors.Add(error);
                    }
                }


                if (person.DateAnciennete != null)
                {
                    if (!dateRegex.IsMatch(person.DateAnciennete))
                    {

                        Error error = new Error()
                        {
                            Code = "ERR-007",
                            Description_EN = "The date of seniority format is invalid",
                            Description_FR = "Le format de la date d'ancienneté est invalide",
                            RecordIndex = row
                        };

                        recordSet.Errors.Add(error);
                    }
                }


                if (person.DateStatut != null)
                    person.DateStatut = "1900-01-01";

                recordSet.Records.Add(person);
            }


            if (recordSet.Records.Count(x => x.Secteur == "") == recordSet.Records.Count())
            {

                foreach (var record in recordSet.Records)
                {
                    record.Secteur = "Général";
                }

                Error error = new Error()
                {
                    Code = "ERR-008",
                    Description_EN = "The sector is required",
                    Description_FR = "Le secteur est requis",
                    RecordIndex = 0
                };

                recordSet.Errors.Add(error);
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

    }
}
