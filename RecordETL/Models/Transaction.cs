namespace RecordETL.Models
{
    public class Transaction
    {
        public int Row {get; set;}
        public string NumeroMembre { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string DepositDate { get; set; }
        public string Amount { get; set; }
        public string Type { get; set; }
        public string Account { get; set; }
        public string Note { get; set; }
        public string CompanyCode { get; set; }
        public string HoursWorked { get; set; }
        public string SourceOfPayment { get; set; }
        public string WorkingGrossDues { get; set; }
        public string ControlNo { get; set; }
    }
}
