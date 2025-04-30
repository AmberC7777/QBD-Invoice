using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
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
        public decimal Quantity { get; set; }
        public decimal Rate { get; set; }
        public decimal Amount { get; set; }
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
            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                return csv.GetRecords<InvoiceLine>().ToList();
            }
        }

        static void Main(string[] args)
        {
            var headers = ReadInvoiceHeaders("InvoiceHeader.csv");
            var lines = ReadInvoiceLines("InvoiceLines.csv");

            PushInvoicesToQuickBooks(headers, lines);
        }

        public static void PushInvoicesToQuickBooks(
            List<InvoiceHeader> headers,
            List<InvoiceLine> lines)
        {
            var sessionMgr = new QBSessionManager();
            try
            {
                sessionMgr.OpenConnection2(
                    "",
                    "InvoiceImporter",
                    ENConnectionType.ctLocalQBD
                );
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

                        lineAdd.Quantity.SetValue((double)ln.Quantity);
                        lineAdd.ORRatePriceLevel.Rate.SetValue((double)ln.Rate);
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
