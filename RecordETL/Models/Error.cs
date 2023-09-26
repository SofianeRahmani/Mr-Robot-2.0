namespace RecordETL.Models
{
    public class Error
    {
        public string Code { get; set; }
        public string Description_EN { get; set; }
        public string Description_FR { get; set; }
        public int RecordIndex { get; set; }
    }
}
