namespace RecordETL.Models
{
    public class Record
    {
        public string NumeroMembre { get; set; }
        public string Nom { get; set; }
        public string Prenom { get; set; }
        public string Sexe { get; set; }
        public string CourrielTravail { get; set; }
        public string CourrielPersonnel { get; set; }
        public string CourrielAutre { get; set; }
        public string Telephone { get; set; }
        public string TelephoneTravail { get; set; }
        public string TelephoneCellulaire { get; set; }
        public string Adresse { get; set; }
        public string Ville { get; set; }
        public string Province { get; set; }
        public string CodePostal { get; set; }
        public string Nas { get; set; }
        public string Categories { get; set; }
        public string DateNaissance { get; set; }
        public string DateAnciennete { get; set; }
        public string Anciennete { get; set; }
        public string DateEmbauche { get; set; }
        public string Statut { get; set; }
        public string DateStatut { get; set; }
        public string IdSystemeSource { get; set; }
        
        public string Secteur_FR { get; set; }
        public string Secteur_EN { get; set; }
        public string Acronyme { get; set; }
        
        public string StatutPersonne { get; set; }
        public string IdentifiantAlternatif { get; set; }
        public string InfosComplementaires1 { get; set; }
        public string InfosComplementaires2 { get; set; }
        
        public string Employeur { get; set; }
        public string NumeroEmployeur { get; set; }
        public string Fonction { get; set; }
        public string DateDebut { get; set; }
        public string DateFin { get; set; }
        public string InfoComplementairesEmploi { get; set; }
        
        public string IdPastille { get; set; }
        public string DescriptionPastille { get; set; }
        public string DateEvenement { get; set; }
        public string Note { get; set; }
    }
}
