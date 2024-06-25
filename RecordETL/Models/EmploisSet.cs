namespace RecordETL.Models;
using System.Collections.Generic;

public class EmploisSet
{
    public List<Emplois> Emplois { get; set; }
    
    public List<Error> Errors { get; set; }
}