namespace RecordETL.Models
{
    public class Transaction
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string DepositDate { get; set; }
        public string Amount { get; set; }
        public string Type { get; set; }
        public string Account { get; set; }
        public string Note { get; set; }
    }
}
