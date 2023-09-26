using RecordETL.Models;
using System.Collections.Generic;

namespace RecordETL.Services
{
    public class ValidatorService
    {


        public RecordSet Validate(RecordSet recordSet)
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


            return recordSet;
        }


        private List<Error> ValidateRuleSet1(Record record, int index)
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
                    RecordIndex = index
                };

                errors.Add(error);
            }

            // Rule 2: Required field Nom
            if (string.IsNullOrEmpty(record.Nom))
            {
                var error = new Error()
                {
                    Code = "ERR-002",
                    Description_EN = "Lastname is required",
                    Description_FR = "Nom est requis",
                    RecordIndex = index
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
                    RecordIndex = index
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
                    RecordIndex = index
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
                    RecordIndex = index
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
                    RecordIndex = index
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
                    RecordIndex = index
                };

                errors.Add(error);
            }
            
            return errors;
        }


    }
}
