using System.Collections.Generic;

namespace RecordETL.Models;

public class Emplois
{
    public int Row {get; set;}
    
    public string NumeroMembre { get; set; }
    
    public string Nom { get; set; }
    
    public string Prenom { get; set; }
    
    public string Employeur { get; set; }
    
    public string NumEmployeur { get; set; }
    
    public string Secteur { get; set; }
    
    public string fonction { get; set; }
    
    public string DateDebut { get; set; }
    
    public string DateFin { get; set; }
    
    public Dictionary<string,string> infosComplementaireEmplois { get; set; }
    
}