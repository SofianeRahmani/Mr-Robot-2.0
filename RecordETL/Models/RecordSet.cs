using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RecordETL.Models
{
    public class RecordSet
    {
        public List<Record> Records { get; set; }
        public List<Error> Errors { get; set; }
    }
}
