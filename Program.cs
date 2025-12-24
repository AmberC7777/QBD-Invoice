using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using QBFC16Lib;   // COM reference for QBFC16

namespace QBD_Invoice
{
    public class InvoiceHeader
    {
        public string InvoiceID { get; set; } = string.Empty;
        public string CustomerRef { get; set; } = string.Empty;
        public DateTime TxnDate { get; set; }
        public string RefNumber { get; set; } = string.Empty;
    }

    public class InvoiceLine
    {
        public string InvoiceID { get; set; } = string.Empty;
        public int LineNum { get; set; }
        public string ItemRef { get; set; } = string.Empty;
        public string Desc { get; set; } = string.Empty;
        public decimal? Quantity { get; set; }   // nullable
        public decimal? Rate { get; set; }       // nullable
        public decimal? Amount { get; set; }     // nullable
    }

    class Program
    {
        static List<InvoiceHeader> ReadInvoiceHeaders(string path)
        {
            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                return csv.GetRecords<InvoiceHeader>().ToList();
            }
        }

        static List<InvoiceLine> ReadInvoiceLines(string path)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                MissingFieldFound = null,
                HeaderValidated = null,
            };

            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, config))
            {
                // Treat empty strings as null for decimal?
                csv.Context.TypeConverterOptionsCache
                   .GetOptions<decimal?>()
                   .NullValues.Add(string.Empty);

                return csv.GetRecords<InvoiceLine>().ToList();
            }
        }

        static void Main()
        {
            try
            {
                var headers = ReadInvoiceHeaders("InvoiceHeader.csv");
                var lines = ReadInvoiceLines("InvoiceLines.csv");
                PushInvoicesToQuickBooks(headers, lines);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Unhandled exception:");
                Console.WriteLine(ex);
                Console.ResetColor();
            }
            finally
            {
                Console.WriteLine();
                Console.WriteLine("Execution complete. Press any key to exit...");
                Console.ReadKey();
            }
        }

        public static void PushInvoicesToQuickBooks(
            List<InvoiceHeader> headers,
            List<InvoiceLine> lines)
        {
            var sessionMgr = new QBSessionManager();
            try
            {
                sessionMgr.OpenConnection2("", "InvoiceImporter", ENConnectionType.ctLocalQBD);
                sessionMgr.BeginSession("", ENOpenMode.omDontCare);

                var msgSet = sessionMgr.CreateMsgSetRequest("US", 16, 0);
                msgSet.Attributes.OnError = ENRqOnError.roeContinue;

                var invoiceOrder = new List<string>();

                foreach (var hdr in headers)
                {
                    var invRq = msgSet.AppendInvoiceAddRq();
                    invoiceOrder.Add(hdr.InvoiceID);

                    invRq.CustomerRef.FullName.SetValue(hdr.CustomerRef);
                    invRq.TxnDate.SetValue(hdr.TxnDate);
                    invRq.RefNumber.SetValue(hdr.RefNumber);

                    foreach (var ln in lines.Where(l => l.InvoiceID == hdr.InvoiceID))
                    {
                        var lineAdd = invRq
                            .ORInvoiceLineAddList
                            .Append()
                            .InvoiceLineAdd;

                        lineAdd.ItemRef.FullName.SetValue(ln.ItemRef);
                        if (!string.IsNullOrWhiteSpace(ln.Desc))
                            lineAdd.Desc.SetValue(ln.Desc);

                        // If this is your percent-based item ("OOP"), send Rate as PERCENTTYPE via RatePercent.
                        // NOTE: PERCENTTYPE is expressed like "5" for 5% (not 0.05).
                        bool isOop = string.Equals(ln.ItemRef, "OOP", StringComparison.OrdinalIgnoreCase);

                        // Only send Quantity if CSV cell had a value AND the line is not a percent-based line.
                        // (Percent discount-style lines typically do not use Quantity.)
                        if (!isOop && ln.Quantity.HasValue)
                        {
                            lineAdd.Quantity.SetValue((double)ln.Quantity.Value);
                        }

                        // Only send Rate/RatePercent if CSV cell had a value
                        if (ln.Rate.HasValue)
                        {
                            if (isOop)
                            {
                                // <-- This is the key change: send as PERCENTTYPE
                                lineAdd.ORRatePriceLevel.RatePercent.SetValue((double)ln.Rate.Value);
                            }
                            else
                            {
                                lineAdd.ORRatePriceLevel.Rate.SetValue((double)ln.Rate.Value);
                            }
                        }

                        // (You can similarly handle Amount if you ever map it.)
                    }
                }

                var respSet = sessionMgr.DoRequests(msgSet);
                var errorLog = new List<string>();

                for (int i = 0; i < respSet.ResponseList.Count; i++)
                {
                    var resp = respSet.ResponseList.GetAt(i);
                    var invoiceID = invoiceOrder[i];

                    if (resp.StatusCode == 0)
                    {
                        var invRet = (IInvoiceRet)resp.Detail;
                        Console.WriteLine(
                          $"✔ Invoice {invoiceID} created: TxnID={invRet.TxnID.GetValue()}"
                        );
                    }
                    else
                    {
                        var msg =
                          $"✘ Invoice {invoiceID} failed: " +
                          $"Code={resp.StatusCode} – {resp.StatusMessage}";
                        Console.WriteLine(msg);
                        errorLog.Add(msg);
                    }
                }

                if (errorLog.Any())
                {
                    File.WriteAllLines("InvoiceImportErrors.log", errorLog);
                    Console.WriteLine("Errors logged to InvoiceImportErrors.log");
                }
            }
            finally
            {
                sessionMgr.EndSession();
                sessionMgr.CloseConnection();
            }
        }
    }
}
