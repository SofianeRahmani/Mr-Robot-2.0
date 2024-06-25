namespace RecordETL.Models;
using System.Collections.Generic;
public class EvenementSet
{
    public List<Evenement> Evenements { get; set; }
    
    public List<Error> Errors { get; set; }
}