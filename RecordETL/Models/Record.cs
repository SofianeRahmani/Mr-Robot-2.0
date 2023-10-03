using System.Collections.Generic;

namespace RecordETL.Models
{
    public class Record
    {
        public int Row {get; set;}

        public string NumeroMembre { get; set; }
        public string Nom { get; set; }
        public string Prenom { get; set; }
        public string SecondNom { get; set; }
        public string SecondPrenom { get; set; }

        public string Sexe { get; set; }
        public string CourrielTravail { get; set; }
        public string CourrielPersonnel { get; set; }
        public string CourrielAutre { get; set; }
        public string Categorie{get; set; }

        public string Telephone { get; set; }
        public string TelephoneTravail { get; set; }
        public string TelephoneCellulaire { get; set; }
        public string NumeroAppartement { get; set; }
        public string Adresse { get; set; }
        public string Ville { get; set; }
        public string Province { get; set; }
        public string Pays { get; set; }
        public string CodePostal { get; set; }
        public string Nas { get; set; }
        public string Categories { get; set; }
        public string DateNaissance { get; set; }
        public string DateAnciennete { get; set; }
        public string Anciennete { get; set; }
        public string Statut { get; set; }
        public string DateStatut { get; set; }
        public string Secteur { get; set; }
        public string StatutPersonne { get; set; }
        public string IdentifiantAlternatif { get; set; }
        public string InfosComplementaires1 { get; set; }
        public string InfosComplementaires2 { get; set; }

        public string Employeur { get; set; }
        public string NumeroEmployeur { get; set; }
        public string Fonction { get; set; }
        public string DateDebut { get; set; }
        public string DateFin { get; set; }
        public string InfosComplementairesEmplois { get; set; }


        public List<Transaction> Transactions { get; set; }

        public Record()
        {
            Transactions = new List<Transaction>();
        }
    }
}
