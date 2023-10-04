using System.Collections.Generic;

namespace RecordETL.Models
{
    public class MembresSet
    {
        public List<Membre> Records { get; set; }
        public List<Error> Errors { get; set; }
    }
}
