using RecordETL.Models;
using System.Collections.Generic;
using System.Linq;

namespace RecordETL.Services
{
    public class ValidatorService
    {
        public static RecordSet Validate(RecordSet recordSet)
        {
            // Validate the recordSet
            // if the recordSet is not valid
            // add an error to the recordSet.Errors list
            for (int index = 0; index < recordSet.Records.Count; index++)
            {
                // Validate using RuleSet 1
                List<Error> errors = ValidateRuleSet1(recordSet.Records[index], index);
                recordSet.Errors.AddRange(errors);
            }
            ValidateNumeroMembre(recordSet);

            return recordSet;
        }


        private static List<Error> ValidateRuleSet1(Record record, int index)
        {
            List<Error> errors = new List<Error>();
            // Rule 1: Required field NumeroMembre
            if (string.IsNullOrEmpty(record.NumeroMembre))
            {
                var error = new Error()
                {
                    Code = "ERR-001",
                    Description_EN = "MemberNumber is required",
                    Description_FR = "NumeroMembre est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            /*
            // Rule 2: Required field Nom
            if (string.IsNullOrEmpty(record.Nom))
            {
                var error = new Error()
                {
                    Code = "ERR-002",
                    Description_EN = "Lastname is required",
                    Description_FR = "Nom est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            // Rule 3: Required field Prenom
            if (string.IsNullOrEmpty(record.Prenom))
            {
                var error = new Error()
                {
                    Code = "ERR-003",
                    Description_EN = "Name is required",
                    Description_FR = "Prenom est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            // Rule 4: Required field Statut
            if (string.IsNullOrEmpty(record.Statut))
            {
                var error = new Error()
                {
                    Code = "ERR-004",
                    Description_EN = "Status is required",
                    Description_FR = "Statut est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            // Rule 5: Required field DateStatut
            if (string.IsNullOrEmpty(record.DateStatut))
            {
                var error = new Error()
                {
                    Code = "ERR-005",
                    Description_EN = "DateStatus is required",
                    Description_FR = "DateStatut est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            // Rule 6: Required field IdSystemeSource
            if (string.IsNullOrEmpty(record.IdSystemeSource))
            {
                var error = new Error()
                {
                    Code = "ERR-006",
                    Description_EN = "IdSystemSource is required",
                    Description_FR = "IdSystemeSource est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }

            // Rule 7: Required field Secteur
            if (string.IsNullOrEmpty(record.Secteur))
            {
                var error = new Error()
                {
                    Code = "ERR-007",
                    Description_EN = "Sector is required",
                    Description_FR = "Secteur est requis",
                    RecordIndex = record.row
                };

                errors.Add(error);
            }
            */

            return errors;
        }


        private static void ValidateNumeroMembre(RecordSet recordSet)
        {

            // replace missing
            var missingNumeroMembre = recordSet.Records.Where(r => string.IsNullOrEmpty(r.NumeroMembre)).ToList();
            for (int i = 0; i < missingNumeroMembre.Count; i++)
            {
                var record = missingNumeroMembre[i];

                int number = i + 1;
                string value = number < 10 ? $"00{number}" : number < 100 ? $"0{number}" : number.ToString();
                record.NumeroMembre = $"SN-{value}";
            }


            // remove duplicates
            var groupedRecords = from r in recordSet.Records
                                 group r by r.NumeroMembre into g
                                 where g.Count() > 1
                                 select g;

            foreach (var group in groupedRecords)
            {
                var records = group.ToList();
                for (int i = 0; i < records.Count; i++)
                {
                    var record = records[i];
                    if (i > 0) // Do not modify the first record
                    {
                        record.NumeroMembre = $"{record.NumeroMembre}-D{i}"; // Append D1, D2, D3, etc. to duplicates

                        var error = new Error()
                        {
                            Code = "ERR-007",
                            Description_EN = "Sector is required",
                            Description_FR = "Secteur est requis",
                            RecordIndex = record.row
                        };

                        recordSet.Errors.Add(error);
                    }
                }
            }


        }
    }
}
