using OfficeOpenXml;
using RecordETL.Models;
using System.Collections.Generic;
using System.IO;

namespace RecordETL.Services
{
    public class ExtractorService
    {

        public static List<string> ReadColumnsNames(string filePath, int pageIndex)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<string> columns = new List<string>();

            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[pageIndex];


            if (worksheet.Dimension == null) return columns; // Return empty list if worksheet is empty


            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                columns.Add(worksheet.Cells[3, col].Text);
            }

            return columns;
        }

        public static RecordSet Extract(string filePath, int pageIndex, List<AttributeIndex> columns)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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

                var NumeroMembre = columns[1].Index != -1 ? worksheet.Cells[row, columns[1].Index +1].Text : null;
                var Nom = columns[2].Index != -1 ?worksheet.Cells[row, columns[2].Index+1].Text : null;
                var Prenom = columns[3].Index != -1 ?worksheet.Cells[row, columns[3].Index+1].Text : null;
                var Sexe = columns[4].Index != -1 ? worksheet.Cells[row, columns[4].Index+1].Text : null;
                var CourrielTravail = columns[5].Index != -1 ? worksheet.Cells[row, columns[5].Index+1].Text : null;
                var CourrielPersonnel = columns[6].Index != -1 ? worksheet.Cells[row, columns[6].Index + 1].Text : null;
                var CourrielAutre = columns[7].Index != -1 ? worksheet.Cells[row, columns[7].Index+1].Text : null;
                var Telephone = columns[8].Index != -1 ? worksheet.Cells[row, columns[8].Index +1].Text : null;
                var TelephoneTravail = columns[9].Index != -1 ? worksheet.Cells[row, columns[9].Index+1].Text : null;
                var TelephoneCellulaire = columns[10].Index != -1 ? worksheet.Cells[row, columns[10].Index + 1].Text : null;
                var Adresse = columns[11].Index != -1 ? worksheet.Cells[row, columns[11].Index +1].Text : null;
                var Ville = columns[12].Index != -1 ? worksheet.Cells[row, columns[12].Index + 1].Text : null;
                var Province = columns[13].Index != -1 ? worksheet.Cells[row, columns[13].Index + 1].Text : null;
                var CodePostal = columns[14].Index != -1 ? worksheet.Cells[row, columns[14].Index +1].Text : null;
                var Nas = columns[15].Index != -1 ? worksheet.Cells[row, columns[15].Index +1].Text : null;
                var Categories = columns[16].Index != -1 ? worksheet.Cells[row, columns[16].Index + 1].Text : null;
                var DateNaissance = columns[17].Index != -1 ? worksheet.Cells[row, columns[17].Index +1].Text : null;
                var DateAnciennete = columns[18].Index != -1 ? worksheet.Cells[row, columns[18].Index + 1].Text : null;
                var Anciennete = columns[19].Index != -1 ? worksheet.Cells[row, columns[19].Index + 1].Text : null;
                var DateEmbauche = columns[20].Index != -1 ? worksheet.Cells[row, columns[20].Index + 1].Text : null;
                var Statut = columns[21].Index != -1 ? worksheet.Cells[row, columns[21].Index + 1].Text : null;
                var DateStatut = columns[22].Index != -1 ? worksheet.Cells[row, columns[22].Index + 1].Text : null;
                var IdSystemeSource = columns[23].Index != -1 ? worksheet.Cells[row, columns[23].Index + 1].Text : null;
                var Secteur = columns[24].Index != -1 ? worksheet.Cells[row, columns[24].Index + 1].Text : null;
                var StatutPersonne = columns[25].Index != -1 ? worksheet.Cells[row, columns[25].Index +1].Text : null;
                var IdentifiantAlternatif = columns[26].Index != -1 ? worksheet.Cells[row, columns[26].Index +1].Text : null;
                var InfosComplementaires1 = columns[27].Index != -1 ? worksheet.Cells[row, columns[27].Index + 1].Text : null;
                var InfosComplementaires2 = columns[28].Index != -1 ? worksheet.Cells[row, columns[28].Index + 1].Text : null;

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
