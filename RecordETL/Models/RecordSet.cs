using System.Collections.Generic;

namespace RecordETL.Models
{
    public class RecordSet
    {
        public List<Record> Records { get; set; }
        public List<Error> Errors { get; set; }
    }
}
