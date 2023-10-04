using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace RecordETL.Services
{
    public class EmployeursService
    {

        public static List<string> ReadColumnsNames(ExcelWorkbook workbook)
        {
            List<string> columns = new List<string>();
            var worksheet = workbook.Worksheets[2];
            if (worksheet.Dimension == null) return columns; // Return empty list if worksheet is empty
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                columns.Add(worksheet.Cells[1, col].Text);
            }

            return columns;
        }



        public static string? GetColumnValue(int row, int column, ExcelWorksheet worksheet)
        {
            return column != -1 ? worksheet.Cells[row, column + 1].Text.Trim() : null;
        }



        public static TransactionsSet ReadAndValidate(ExcelWorkbook workbook,
            List<AttributeIndex> positions)
        {

            TransactionsSet Set = new TransactionsSet();
            Set.Transactions = new List<Transaction>();
            Set.Errors = new List<Models.Error>();

            var sheet = workbook.Worksheets[2];


            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                bool empty = true;
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    if (sheet.Cells[row, col].Text != "")
                    {
                        empty = false;
                    }
                }

                if (empty) break;

                var transaction = new Transaction();
                Type type = typeof(Transaction);
                transaction.Row = row;

                foreach (var position in positions)
                {
                    PropertyInfo propertyInfo = type.GetProperty(position.Name);
                    propertyInfo.SetValue(transaction, GetColumnValue(row, position.Index, sheet));
                }



                if (transaction.StartDate != null)
                {
                    var dateNaissance = FormatDate(transaction.StartDate);

                    if (dateNaissance == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-006",
                            Description_EN = "The date of start format is invalid",
                            Description_FR = "Le format de la date de debut est invalide",
                            RecordIndex = row
                        };

                        Set.Errors.Add(error);
                    }
                    else
                    {
                        transaction.StartDate = dateNaissance.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    transaction.StartDate = "Inconnue";
                }


                if (transaction.EndDate != null)
                {
                    var dateAnciennete = FormatDate(transaction.EndDate);

                    if (dateAnciennete == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-007",
                            Description_EN = "The date of end format is invalid",
                            Description_FR = "Le format de la date de fin est invalide",
                            RecordIndex = row
                        };

                        Set.Errors.Add(error);
                    }
                    else
                    {
                        transaction.EndDate = dateAnciennete.Value.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    transaction.EndDate = "Inconnue";
                }


                if (transaction.DepositDate != null)
                {
                    var dateStatut = FormatDate(transaction.DepositDate);

                    if (dateStatut == null)
                    {
                        Error error = new Error()
                        {
                            Code = "ERR-008",
                            Description_EN = "The date of Deposit format is invalid",
                            Description_FR = "Le format de la date de depot est invalide",
                            RecordIndex = row
                        };

                        Set.Errors.Add(error);
                    }
                    else
                    {
                        transaction.DepositDate = dateStatut.Value.ToString("yyyy-MM-dd");
                    }
                }else {
                    transaction.DepositDate = "Inconnue";
                }
                
                Set.Transactions.Add(transaction);
            }

            return Set;
        }



        static DateTime? FormatDate(string date)
        {
            date = date.Trim();
            date = date.Replace("'", "");

            string[] formats = { "M/d/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "d MMMM yyyy", "dd MMMMM yyyy", "dd-MMM-yy" };
            DateTime dt;
            if (DateTime.TryParseExact(date, formats, new CultureInfo("fr-FR"), DateTimeStyles.None, out dt))
            {
                return dt;
            }

            if (DateTime.TryParseExact(date, formats, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
            {
                return dt;
            }

            return null;
        }

    }
}
