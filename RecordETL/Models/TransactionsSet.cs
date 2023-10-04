using System.Collections.Generic;

namespace RecordETL.Models
{
    public class TransactionsSet
    {

        public List<Transaction> Transactions { get; set; }
        public List<Error> Errors { get; set; }
    }
}
