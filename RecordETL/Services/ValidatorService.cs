using RecordETL.Models;
using System.Collections.Generic;
using System.Linq;

namespace RecordETL.Services
{
    public class ValidatorService
    {
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
    }
}
