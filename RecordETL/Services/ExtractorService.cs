using OfficeOpenXml;
using RecordETL.Models;
using System.Collections.Generic;
using System.IO;

namespace RecordETL.Services
{
    public class ExtractorService
    {


        public RecordSet Extract(string filePath)
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
            var worksheet = workbook.Worksheets[0];


            // Assuming that the first row contains headers
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {

                //
                var name = worksheet.Cells[row, 1].Text; // Assuming Name is in the first column
                var age = int.Parse(worksheet.Cells[row, 2].Text); // Assuming Age is in the second column

                var person = new Record()
                {
                    
                };

                recordSet.Records.Add(person);
            }




            return recordSet;
        }
    }
}
