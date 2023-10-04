using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace RecordETL.Services
{
    public class ExtractorService
    {

        public static List<string> ReadDataSourceColumnsNames(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<string> columns = new List<string>();

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[1];
            
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



        public static MembresSet ExtractMembres(string filePath, List<AttributeIndex> positions, bool isAmerican, string terminaisonCourriel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            MembresSet membresSet = new MembresSet();
            membresSet.Records = new List<Membre>();
            membresSet.Errors = new List<Models.Error>();

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;


            var datasourceSheet = workbook.Worksheets[1];



            List<string> fonctions = new List<string>();
            var fonctionsSheet = workbook.Worksheets[7];

            for (int row = 2; row <= fonctionsSheet.Dimension.End.Row; row++)
            {
                string value = GetColumnValue(row, 0, fonctionsSheet);

                if (value == null || value == "") break;

                fonctions.Add(value);
            }

            List<Tuple<string, string>> secteurs = new List<Tuple<string, string>>();
            var secteursSheet = workbook.Worksheets[8];
            for (int row = 2; row <= secteursSheet.Dimension.End.Row; row++)
            {
                string secteur_fr = GetColumnValue(row, 0, secteursSheet);
                string secteur_en = GetColumnValue(row, 1, secteursSheet);

                if ((secteur_fr == null || secteur_fr == "") && (secteur_en == null || secteur_en == "")) break;

                secteurs.Add(new Tuple<string, string>(secteur_fr, secteur_en));
            }

            for (int row = 2; row <= datasourceSheet.Dimension.End.Row; row++)
            {
                bool empty = true;
                for (int col = 1; col <= datasourceSheet.Dimension.End.Column; col++)
                {
                    if (datasourceSheet.Cells[row, col].Text != "")
                    {
                        empty = false;
                    }
                }

                if (empty) break;

                var person = new Membre();
                Type type = typeof(Membre);
                person.Row = row;

                foreach (var position in positions)
                {
                    PropertyInfo propertyInfo = type.GetProperty(position.Name);
                    propertyInfo.SetValue(person, GetColumnValue(row, position.Index, datasourceSheet));
                }

                person.Nom = person.Nom?.Trim() + " " + person.SecondNom;
                person.Prenom = person.Prenom?.Trim() + " " + person.SecondPrenom;

                person.Telephone = FormatPhoneNumber(person.Telephone);
                if (person.Telephone != null && person.Telephone.Length != 10)
                {
                    Error error = new Error()
                    {
                        Code = "ERR-002",
                        Description_EN = "The phone number must be composed of 10 digits.",
                        Description_FR = "Le numéro de téléphone doit être composé de 10 chiffres.",
                        RecordIndex = row
                    };

                    membresSet.Errors.Add(error);
                }

                person.TelephoneTravail = FormatPhoneNumber(person.TelephoneTravail);
                if (person.Telephone != null && person.Telephone.Length != 10)
                {
                    Error error = new Error()
                    {
                        Code = "ERR-003",
                        Description_EN = "The work phone number must be composed of 10 digits.",
                        Description_FR = "Le numéro de téléphone du travail doit être composé de 10 chiffres.",
                        RecordIndex = row
                    };
                    membresSet.Errors.Add(error);
                }

                person.TelephoneCellulaire = FormatPhoneNumber(person.TelephoneCellulaire);

                if (person.TelephoneCellulaire != null && person.TelephoneCellulaire.Length != 10)
                {
                    Error error = new Error()
                    {
                        Code = "ERR-004",
                        Description_EN = "The cell phone number must be composed of 10 digits.",
                        Description_FR = "Le numéro de téléphone cellulaire doit être composé de 10 chiffres.",
                        RecordIndex = row
                    };
                    membresSet.Errors.Add(error);
                }


                if (terminaisonCourriel == null)
                {
                    person.CourrielTravail = person.CourrielTravail?.Trim();
                    person.CourrielPersonnel = person.CourrielPersonnel?.Trim();
                    person.CourrielAutre = person.CourrielAutre?.Trim();
                }
                else
                {

                    string? domain = person.CourrielTravail?.Split('@')[1];

                    if (domain != null && !domain.Contains(terminaisonCourriel))
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

                        membresSet.Errors.Add(error);
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

                            membresSet.Errors.Add(error);
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

                            membresSet.Errors.Add(error);
                        }
                        else
                        {
                            person.CodePostal = $"{person.CodePostal.Substring(0, 3)} {person.CodePostal.Substring(3, 3)}";
                        }
                    }
                }

                
                if (person.DateNaissance != null)
                {
                    var dateNaissance = FormatDate(person.DateNaissance);

                    if (dateNaissance == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-006",
                            Description_EN = "The date of birth format is invalid",
                            Description_FR = "Le format de la date de naissance est invalide",
                            RecordIndex = row
                        };

                        membresSet.Errors.Add(error);
                    }
                    else
                    {
                        person.DateNaissance = dateNaissance.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    person.DateNaissance = "Inconnue";
                }


                if (person.DateAnciennete != null)
                {
                    var dateAnciennete = FormatDate(person.DateAnciennete);

                    if (dateAnciennete == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-007",
                            Description_EN = "The date of seniority format is invalid",
                            Description_FR = "Le format de la date d'ancienneté est invalide",
                            RecordIndex = row
                        };

                        membresSet.Errors.Add(error);
                    }
                    else
                    {
                        person.DateAnciennete = dateAnciennete.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    person.DateAnciennete = "Inconnue";
                }


                if (person.DateStatut != null)
                {
                    var dateStatut = FormatDate(person.DateStatut);

                    if (dateStatut == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-008",
                            Description_EN = "The status date format is invalid",
                            Description_FR = "Le format de la date de statut est invalide",
                            RecordIndex = row
                        };

                        membresSet.Errors.Add(error);
                    }
                    else
                    {
                        person.DateStatut = dateStatut.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    person.DateStatut = "1900-01-01";
                }


                if (person.DateDebut != null)
                {
                    var dateDebutMandat = FormatDate(person.DateDebut);

                    if (dateDebutMandat == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-009",
                            Description_EN = "The mandate start date format is invalid",
                            Description_FR = "Le format de la date de début de mandat est invalide",
                            RecordIndex = row
                        };

                        membresSet.Errors.Add(error);
                    }
                    else
                    {
                        person.DateDebut = dateDebutMandat.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    person.DateDebut = "Inconnue";
                }



                if (person.Fonction != null && fonctions.Count() > 0)
                {
                    if (!fonctions.Contains(person.Fonction))
                    {
                        person.Fonction = "";

                        Error error = new Error()
                        {
                            Code = "ERR-008",
                            Description_EN = "The function is invalid",
                            Description_FR = "La fonction est invalide",
                            RecordIndex = row
                        };

                        membresSet.Errors.Add(error);
                    }
                }


                if (person.Secteur == null || person.Secteur == "")
                {
                    person.Secteur = "Général";
                }
                else
                {
                    if (secteurs.Any())
                    {

                        bool exist = false;
                        foreach (var secteur in secteurs)
                        {
                            if (secteur.Item1 == person.Secteur || secteur.Item2 == person.Secteur)
                            {
                                exist = true;
                                break;
                            }
                        }

                        if (!exist)
                        {
                            Error error = new Error()
                            {
                                Code = "ERR-009",
                                Description_EN = "The sector is invalid",
                                Description_FR = "Le secteur est invalide",
                                RecordIndex = row
                            };

                            membresSet.Errors.Add(error);
                        }
                    }
                }

                membresSet.Records.Add(person);
            }



            return membresSet;
        }


        static string FormatPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
            {
                return phoneNumber;
            }

            // remove all non-numeric characters
            phoneNumber = Regex.Replace(phoneNumber, @"[^0-9]", "");

            return phoneNumber;
        }

        static DateTime? FormatDate(string date)
        {
            date = date.Trim();
            date = date.Replace("'", "");

            string[] formats = { "M/d/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "d MMMM yyyy", "dd MMMMM yyyy" };
            DateTime dt;
            if (DateTime.TryParseExact(date, formats, new CultureInfo("fr-FR"), DateTimeStyles.None, out dt))
            {
                return dt;
            }
            
            if(DateTime.TryParseExact(date, formats, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
            {
                return dt;
            }

            return null;
        }

    }
}
