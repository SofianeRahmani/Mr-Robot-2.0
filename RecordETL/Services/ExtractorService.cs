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

        public static RecordSet Extract(string filePath, int pageIndex, List<AttributeIndex> positions)
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

                var NumeroMembre = positions[0].Index != -1 ? worksheet.Cells[row, positions[0].Index].Text : null;

                var Nom = positions[1].Index != -1 ? worksheet.Cells[row, positions[1].Index+1].Text : null;
                var Prenom = positions[2].Index != -1 ? worksheet.Cells[row, positions[2].Index+1].Text : null;
                var Sexe = positions[3].Index != -1 ? worksheet.Cells[row, positions[3].Index+1].Text : null;
                var CourrielTravail = positions[4].Index != -1 ? worksheet.Cells[row, positions[4].Index+1].Text : null;
                var CourrielPersonnel = positions[5].Index != -1 ? worksheet.Cells[row, positions[5].Index + 1].Text : null;
                var CourrielAutre = positions[6].Index != -1 ? worksheet.Cells[row, positions[6].Index+1].Text : null;
                var Telephone = positions[7].Index != -1 ? worksheet.Cells[row, positions[7].Index +1].Text : null;
                var TelephoneTravail = positions[8].Index != -1 ? worksheet.Cells[row, positions[8].Index+1].Text : null;
                var TelephoneCellulaire = positions[9].Index != -1 ? worksheet.Cells[row, positions[9].Index + 1].Text : null;
                var Adresse = positions[10].Index != -1 ? worksheet.Cells[row, positions[10].Index +1].Text : null;
                var Ville = positions[11].Index != -1 ? worksheet.Cells[row, positions[11].Index + 1].Text : null;
                var Province = positions[12].Index != -1 ? worksheet.Cells[row, positions[12].Index + 1].Text : null;
                var CodePostal = positions[13].Index != -1 ? worksheet.Cells[row, positions[13].Index +1].Text : null;
                var Nas = positions[14].Index != -1 ? worksheet.Cells[row, positions[14].Index +1].Text : null;
                var Categories = positions[15].Index != -1 ? worksheet.Cells[row, positions[15].Index + 1].Text : null;
                var DateNaissance = positions[16].Index != -1 ? worksheet.Cells[row, positions[16].Index +1].Text : null;
                var DateAnciennete = positions[17].Index != -1 ? worksheet.Cells[row, positions[17].Index + 1].Text : null;
                var Anciennete = positions[18].Index != -1 ? worksheet.Cells[row, positions[18].Index + 1].Text : null;
                var DateEmbauche = positions[19].Index != -1 ? worksheet.Cells[row, positions[19].Index + 1].Text : null;
                var Statut = positions[20].Index != -1 ? worksheet.Cells[row, positions[20].Index + 1].Text : null;
                var DateStatut = positions[21].Index != -1 ? worksheet.Cells[row, positions[21].Index + 1].Text : null;
                var IdSystemeSource = positions[22].Index != -1 ? worksheet.Cells[row, positions[22].Index + 1].Text : null;
                var Secteur = positions[23].Index != -1 ? worksheet.Cells[row, positions[23].Index + 1].Text : null;
                var StatutPersonne = positions[24].Index != -1 ? worksheet.Cells[row, positions[24].Index +1].Text : null;
                var IdentifiantAlternatif = positions[25].Index != -1 ? worksheet.Cells[row, positions[25].Index +1].Text : null;
                var InfosComplementaires1 = positions[26].Index != -1 ? worksheet.Cells[row, positions[26].Index + 1].Text : null;
                var InfosComplementaires2 = positions[27].Index != -1 ? worksheet.Cells[row, positions[27].Index + 1].Text : null;

                var Employeur = positions[28].Index != -1 ? worksheet.Cells[row, positions[28].Index + 1].Text : null;
                var NumeroEmployeur = positions[29].Index != -1 ? worksheet.Cells[row, positions[29].Index + 1].Text : null;
                var Fonction = positions[30].Index != -1 ? worksheet.Cells[row, positions[30].Index + 1].Text : null;
                var DateDebut = positions[31].Index != -1 ? worksheet.Cells[row, positions[31].Index + 1].Text : null;
                var DateFin = positions[32].Index != -1 ? worksheet.Cells[row, positions[32].Index + 1].Text : null;
                var InfosComplementairesEmplois = positions[33].Index != -1 ? worksheet.Cells[row, positions[33].Index + 1].Text : null;



                var person = new Record()
                {
                    row = row,
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
                    InfosComplementaires2 = InfosComplementaires2,

                    Employeur = Employeur,
                    NumeroEmployeur = NumeroEmployeur,
                    Fonction = Fonction,
                    DateDebut = DateDebut,
                    DateFin = DateFin,
                    InfosComplementairesEmplois = InfosComplementairesEmplois
                };

                recordSet.Records.Add(person);
            }



            return recordSet;
        }
    }
}
