using OfficeOpenXml;
using RecordETL.Models;
using System.Collections.Generic;
using System.IO;

namespace RecordETL.Services
{
    public class ExtractorService
    {


        public RecordSet Extract(string filePath, int pageIndex, List<int> columns)
        {
            RecordSet recordSet = new RecordSet();
            recordSet.Records = new List<Record>();
            recordSet.Errors = new List<Error>();


            // read the excel file
            // for each row in the excel file
            // create a record
            // add the record to the recordSet.Records list

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[pageIndex];


            for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
            {
                // extract the attributes of Record and sort them by the index
                
                var NumeroMembre = worksheet.Cells[row, 1].Text;
                var Nom = worksheet.Cells[row, 2].Text;
                var Prenom = worksheet.Cells[row, 3].Text;
                var Sexe = worksheet.Cells[row, 4].Text;
                var CourrielTravail = worksheet.Cells[row, 5].Text;
                var CourrielPersonnel = worksheet.Cells[row, 6].Text;
                var CourrielAutre = worksheet.Cells[row, 7].Text;
                var Telephone = worksheet.Cells[row, 8].Text;
                var TelephoneTravail = worksheet.Cells[row, 9].Text;
                var TelephoneCellulaire = worksheet.Cells[row, 10].Text;
                var Adresse = worksheet.Cells[row, 11].Text;
                var Ville = worksheet.Cells[row, 12].Text;
                var Province = worksheet.Cells[row, 13].Text;
                var CodePostal = worksheet.Cells[row, 14].Text;
                var Nas = worksheet.Cells[row, 15].Text;
                var Categories = worksheet.Cells[row, 16].Text;
                var DateNaissance = worksheet.Cells[row, 17].Text;
                var DateAnciennete = worksheet.Cells[row, 18].Text;
                var Anciennete = worksheet.Cells[row, 19].Text;
                var DateEmbauche = worksheet.Cells[row, 20].Text;
                var Statut = worksheet.Cells[row, 21].Text;
                var DateStatut = worksheet.Cells[row, 22].Text;
                var IdSystemeSource = worksheet.Cells[row, 23].Text;
                var Secteur = worksheet.Cells[row, 24].Text;
                var StatutPersonne = worksheet.Cells[row, 25].Text;
                var IdentifiantAlternatif = worksheet.Cells[row, 26].Text;
                var InfosComplementaires1 = worksheet.Cells[row, 27].Text;
                var InfosComplementaires2 = worksheet.Cells[row, 28].Text;

                var person = new Record()
                {
                    NumeroMembre = NumeroMembre,
                    Nom = Nom,
                    Prenom = Prenom,
                    Sexe = Sexe,
                    CourrielTravail = CourrielTravail,
                    CourrielPersonnel = CourrielPersonnel, 
                    CourrielAutre = CourrielAutre,
                    Telephone = Telephone,
                    TelephoneTravail = TelephoneTravail,
                    TelephoneCellulaire = TelephoneCellulaire,
                    Adresse = Adresse,
                    Ville = Ville,
                    Province = Province,
                    CodePostal = CodePostal,
                    Nas = Nas,
                    Categories = Categories,
                    DateNaissance = DateNaissance,
                    DateAnciennete = DateAnciennete,
                    Anciennete = Anciennete,
                    DateEmbauche = DateEmbauche,
                    Statut = Statut,
                    DateStatut = DateStatut,
                    IdSystemeSource = IdSystemeSource,
                    Secteur = Secteur,
                    StatutPersonne = StatutPersonne,
                    IdentifiantAlternatif = IdentifiantAlternatif,
                    InfosComplementaires1 = InfosComplementaires1,
                    InfosComplementaires2 = InfosComplementaires2
                };

                recordSet.Records.Add(person);
            }



            return recordSet;
        }
    }
}
