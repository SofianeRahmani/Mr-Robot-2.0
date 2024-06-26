﻿using OfficeOpenXml;
using RecordETL.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace RecordETL.Services
{
    public class TransactionsService
    {

        public static List<string> ReadColumnsNames(ExcelWorkbook workbook)
        {
            List<string> columns = new List<string>();
            var worksheet = workbook.Worksheets[3];
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

            var sheet = workbook.Worksheets[3];


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
                    var startDate = FormatDate(transaction.StartDate);

                    if (startDate == null)
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
                        transaction.StartDate = startDate.Value.ToString("yyyy-MM-dd");
                    }
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
                }

                Set.Transactions.Add(transaction);
            }

            return Set;
        }



        static DateTime? FormatDate(string date)
        {
            date = date.Trim();
            date = date.Replace("'", "");

            string[] formats = { "d/M/yy", "dd/MM/yyyy", "MM/dd/yyyy", "M/d/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "d MMMM yyyy", "dd MMMMM yyyy", "dd-MMM-yy" };
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

        internal static void ExportCSV(TransactionsSet transactionsSet, string outputPath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{outputPath}\\transactions.csv"))
            {
                file.WriteLine("NumeroMembre,StartDate,EndDate,DepositDate,Amount,Type,Account,Note,CompanyCode,HoursWorked,SourceOfPayment,WorkingGrossDues,ControlNo");
                foreach (var record in transactionsSet.Transactions)
                {
                    file.WriteLine($"{record.NumeroMembre},{record.StartDate},{record.EndDate},{record.DepositDate},{record.Amount},{record.Type},{record.Account},{record.Note},{record.CompanyCode},{record.HoursWorked},{record.SourceOfPayment},{record.WorkingGrossDues},{record.ControlNo}");
                }
            }
        }



        // exports the MembresSet into membres.xlsx
        internal static void ExportErrors(TransactionsSet transactionsSet, ExcelWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add("Transactions Errors");
            worksheet.Cells["A1"].Value = "Code";
            worksheet.Cells["B1"].Value = "Description EN";
            worksheet.Cells["C1"].Value = "Description FR";
            worksheet.Cells["D1"].Value = "Record Index";
            int index = 2;
            foreach (var error in transactionsSet.Errors)
            {
                worksheet.Cells[$"A{index}"].Value = error.Code;
                worksheet.Cells[$"B{index}"].Value = error.Description_EN;
                worksheet.Cells[$"C{index}"].Value = error.Description_FR;
                worksheet.Cells[$"D{index}"].Value = error.RecordIndex;
                index++;
            }
        }
    }
}
