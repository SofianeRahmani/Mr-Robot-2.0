using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace RecordETL.Services
{
    public class MembresService
    {

        public static List<string> ReadColumnsNames(ExcelWorkbook workbook)
        {
            List<string> columns = new List<string>();
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



        public static MembresSet ReadAndValidate(ExcelWorkbook workbook,
            List<AttributeIndex> positions, bool isAmerican, string terminaisonCourriel)
        {

            MembresSet membresSet = new MembresSet();
            membresSet.Records = new List<Membre>();
            membresSet.Errors = new List<Models.Error>();

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

                var membre = new Membre();
                Type type = typeof(Membre);
                membre.Row = row;

                foreach (var position in positions)
                {
                    PropertyInfo propertyInfo = type.GetProperty(position.Name);
                    propertyInfo.SetValue(membre, GetColumnValue(row, position.Index, datasourceSheet));
                }

                membre.Nom = membre.Nom?.Trim() + " " + membre.SecondNom;
                membre.Prenom = membre.Prenom?.Trim() + " " + membre.SecondPrenom;

                membre.Telephone = FormatPhoneNumber(membre.Telephone);
                if (membre.Telephone != null && membre.Telephone.Length != 10)
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

                membre.TelephoneTravail = FormatPhoneNumber(membre.TelephoneTravail);
                if (membre.Telephone != null && membre.Telephone.Length != 10)
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

                membre.TelephoneCellulaire = FormatPhoneNumber(membre.TelephoneCellulaire);

                if (membre.TelephoneCellulaire != null && membre.TelephoneCellulaire.Length != 10)
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
                    membre.CourrielTravail = membre.CourrielTravail?.Trim();
                    membre.CourrielPersonnel = membre.CourrielPersonnel?.Trim();
                    membre.CourrielAutre = membre.CourrielAutre?.Trim();
                }
                else
                {
                    string? domain = membre.CourrielTravail?.Split('@')[1];

                    if (domain != null && !domain.Contains(terminaisonCourriel))
                    {
                        membre.CourrielPersonnel = membre.CourrielTravail;
                        membre.CourrielTravail = null;


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


                if (membre.NumeroAppartement != null)
                {
                    membre.Adresse = membre.NumeroAppartement.Replace("#", "") + " " + membre.Adresse?.Trim();
                }




                if (membre.CodePostal != null)
                {
                    if (isAmerican)
                    {
                        if (membre.CodePostal.Length != 5 || !Regex.IsMatch(membre.CodePostal, @"^\d{5}$"))
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
                        if (membre.CodePostal.Length != 6 || !Regex.IsMatch(membre.CodePostal, @"^[A-Za-z]\d[A-Za-z] \d[A-Za-z]\d$"))
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
                            membre.CodePostal = $"{membre.CodePostal.Substring(0, 3)} {membre.CodePostal.Substring(3, 3)}";
                        }
                    }
                }


                if (membre.DateNaissance != null)
                {
                    var dateNaissance = FormatDate(membre.DateNaissance);

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
                        membre.DateNaissance = dateNaissance.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    membre.DateNaissance = "Inconnue";
                }


                if (membre.DateAnciennete != null)
                {
                    var dateAnciennete = FormatDate(membre.DateAnciennete);

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
                        membre.DateAnciennete = dateAnciennete.Value.ToString("yyyy-MM-dd");
                    }
                }


                if (membre.DateStatut != null)
                {
                    var dateStatut = FormatDate(membre.DateStatut);

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
                        membre.DateStatut = dateStatut.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    membre.DateStatut = "1900-01-01";
                }


                if (membre.DateDebut != null)
                {
                    var dateDebutMandat = FormatDate(membre.DateDebut);

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
                        membre.DateDebut = dateDebutMandat.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    membre.DateDebut = "1900-01-01";
                }



                if (membre.Fonction != null && fonctions.Count() > 0)
                {
                    if (!fonctions.Contains(membre.Fonction))
                    {
                        membre.Fonction = "";

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
                else
                {

                }


                if (membre.Secteur == null || membre.Secteur == null)
                {
                    membre.Secteur = "Général";
                }
                else
                {
                    if (secteurs.Any())
                    {
                        bool exist = false;
                        foreach (var secteur in secteurs)
                        {
                            if (secteur.Item1 == membre.Secteur || secteur.Item2 == membre.Secteur)
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

                membresSet.Records.Add(membre);
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

            string[] formats = { "M/d/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "d MMMM yyyy", "dd MMMMM yyyy", "dd-MMM-yy" };
            DateTime dt;
            if (DateTime.TryParseExact(date, formats, new CultureInfo("fr-FR"), DateTimeStyles.None, out dt))
            {
                return dt;
            }

            if (DateTime.TryParseExact(date, formats, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
            {
                return dt;
            }

            return null;
        }


        public static MembresSet Validate(MembresSet membresSet)
        {
            for (int index = 0; index < membresSet.Records.Count; index++)
            {
                // Validate using RuleSet 1
                List<Error> errors = ValidateRuleSet1(membresSet.Records[index], index);
                membresSet.Errors.AddRange(errors);
            }
            ValidateNumeroMembre(membresSet);

            return membresSet;
        }


        private static List<Error> ValidateRuleSet1(Membre membre, int index)
        {
            List<Error> errors = new List<Error>();
            if (string.IsNullOrEmpty(membre.NumeroMembre))
            {
                var error = new Error()
                {
                    Code = "ERR-001",
                    Description_EN = "MemberNumber is required",
                    Description_FR = "NumeroMembre est requis",
                    RecordIndex = membre.Row
                };

                errors.Add(error);
            }
            return errors;
        }


        private static void ValidateNumeroMembre(MembresSet membresSet)
        {

            // replace missing
            var missingNumeroMembre = membresSet.Records.Where(r => string.IsNullOrEmpty(r.NumeroMembre)).ToList();
            for (int i = 0; i < missingNumeroMembre.Count; i++)
            {
                var record = missingNumeroMembre[i];

                int number = i + 1;
                string value = number < 10 ? $"00{number}" : number < 100 ? $"0{number}" : number.ToString();
                record.NumeroMembre = $"SN-{value}";
                record.Categorie = "A valider- Sans numéro de membre";
            }


            // remove duplicates
            var groupedRecords = from r in membresSet.Records
                                 group r by r.NumeroMembre into g
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
                        record.NumeroMembre = $"{record.NumeroMembre}-D{i}"; // Append D1, D2, D3, etc. to duplicates
                        record.Categorie = " A valider- doublon de numéro de membre";

                        var error = new Error()
                        {
                            Code = "ERR-007",
                            Description_EN = "Sector is required",
                            Description_FR = "Secteur est requis",
                            RecordIndex = record.Row
                        };

                        membresSet.Errors.Add(error);
                    }
                }
            }


        }


        // exports the MembresSet into membres.csv
        internal static void ExportCSV(MembresSet membresSet, string outputPath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{outputPath}\\membres.csv"))
            {
                file.WriteLine("NumeroMembre,Nom,Prenom,DateNaissance,Telephone,TelephoneTravail,TelephoneCellulaire,CourrielTravail,CourrielPersonnel,CourrielAutre,Adresse,Ville,Province,CodePostal,DateAnciennete,DateStatut,DateDebut,Fonction,Secteur,Categorie");
                foreach (var record in membresSet.Records)
                {
                    file.WriteLine($"{record.NumeroMembre},{record.Nom},{record.Prenom},{record.DateNaissance},{record.Telephone},{record.TelephoneTravail},{record.TelephoneCellulaire},{record.CourrielTravail},{record.CourrielPersonnel},{record.CourrielAutre},{record.Adresse},{record.Ville},{record.Province},{record.CodePostal},{record.DateAnciennete},{record.DateStatut},{record.DateDebut},{record.Fonction},{record.Secteur},{record.Categorie}");
                }
            }
        }


        // exports the MembresSet into membres.xlsx
        internal static void ExportErrors(MembresSet membresSet, ExcelWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add("Membres Errors");
            worksheet.Cells["A1"].Value = "Code";
            worksheet.Cells["B1"].Value = "Description EN";
            worksheet.Cells["C1"].Value = "Description FR";
            worksheet.Cells["D1"].Value = "Record Index";
            int index = 2;
            foreach (var error in membresSet.Errors)
            {
                worksheet.Cells[$"A{index}"].Value = error.Code;
                worksheet.Cells[$"B{index}"].Value = error.Description_EN;
                worksheet.Cells[$"C{index}"].Value = error.Description_FR;
                worksheet.Cells[$"D{index}"].Value = error.RecordIndex;
                index++;
            }
        }
    }
}
