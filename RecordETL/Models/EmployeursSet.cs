using System.Collections.Generic;

namespace RecordETL.Models
{
    public class EmployeursSet
    {
        public List<Employeur> Employeurs { get; set; }
        public List<Error> Errors { get; set; }
    }
}
