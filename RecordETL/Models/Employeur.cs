namespace RecordETL.Models
{
    public class Employeur
    {
        public int Row { get; set; }

        public string Numero { get; set; }
        public string Nom { get; set; }
        public string Adresse { get; set; }
        public string Ville { get; set; }
        public string Province { get; set; }
        public string CodePostal { get; set; }

        public string InformationComplémentaire { get; set; }

        public string Telephone { get; set; }
    }
}
