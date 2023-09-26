using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RecordETL.Models;

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
                    Description = "NumeroMembre is required",
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
                    Description = "Nom is required",
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
                    Description = "Prenom is required",
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
                    Description = "Statut is required",
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
                    Description = "DateStatut is required",
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
                    Description = "IdSystemeSource is required",
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
                    Description = "Secteur is required",
                    RecordIndex = index
                };

                errors.Add(error);
            }


            return errors;
        }


    }
}
