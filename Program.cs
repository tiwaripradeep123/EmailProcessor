using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OpenPop.Mime;
using OpenPop.Pop3;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;

namespace Pdf2Text
{
	class Program
	{
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            var sourceBaseFilePath = @"C:\Pradeep\Bills\Download\2_10_2019";
            var cutoffDate = DateTime.Today.AddDays(-101);
            DownloadAttachments(sourceBaseFilePath, cutoffDate);
            bool status = true; 

            if (status)
            {
                var destFilePath = @"C:\Pradeep\Bills\Sorted";
                Regex regexDateTime = new Regex(@"[\d]{4}\s*[\d]{5}\s*[\d]{5}\s*(?<date>[0-9/]+)\s+(?<time>[0-9:]+)");
                Regex regexJobName = new Regex(@"THIS\s*RECEIPT\s*PO\/JOB\s*NAME\:\s*(?<jobname>.*)");

                DirectoryInfo directoryInfo = new DirectoryInfo(sourceBaseFilePath);
                var files = directoryInfo.GetFiles("*.pdf").OrderBy(f => f.LastWriteTime).ToList();
                
                foreach (var file in files)
                {
                    Console.WriteLine($"Processing file {file.Name}");
                    var content = parseUsingPDFBox(file.FullName);
                    var matchDateTime = regexDateTime.Match(content);
                    if (matchDateTime.Success)
                    {
                        var matchJobName = regexJobName.Match(content);
                        string jobName = "-";
                        if (matchJobName.Success)
                        {
                            jobName = matchJobName.Groups["jobname"].Value.Trim();
                        }
                        var date = matchDateTime.Groups["date"].Value.Trim().Replace("/", "-");
                        var time = matchDateTime.Groups["time"].Value.Trim().Replace(":", "-");
                        var fileName = $"{jobName}-{date}-{time}";
                        var destination = Path.Combine(destFilePath, $"{fileName}.pdf");
                        File.Copy(file.FullName, destination, true);
                    }

                }
                
            }
            Console.WriteLine("Processing completed..");
            Console.ReadLine();
        }

        private static bool DownloadAttachments(string sourceBaseFilePath, DateTime cutoffTime)
        {
            try
            {
                var userId = "";
                var password = "";
                string hdsenderName = "HomeDepotReceipt@homedepot.com";
                Pop3Client pop3Client = new Pop3Client();
                pop3Client.Connect("outlook.office365.com", 995, true);
                pop3Client.Authenticate(userId, password);

                int count = pop3Client.GetMessageCount();

                for (int index = count; index > 0; index--)
                {
                    Message messages = null;
                    try
                    {
                        messages = pop3Client.GetMessage(index);
                        var from = messages.Headers.From.Address;
                        if (messages.Headers.DateSent < cutoffTime)
                        {
                            break;
                        }
                        Console.WriteLine($"Processing [{index}] -> [{from}] -> [{messages.Headers.DateSent.ToLongDateString()}]");
                        if (string.Compare(hdsenderName, from, true) == 0)
                        {
                            Console.WriteLine($"Found mail from HD [{messages.Headers.DateSent.ToLongDateString()}]");
                            //string messageText = messages.ToMailMessage().Body;
                            var attachments = messages.FindAllAttachments().Where(x => x.FileName.ToLower().Contains("ereceipt")).ToList();
                            Console.WriteLine($"Found [{attachments.Count}] attachements");
                            var filename = messages.Headers.DateSent.ToLongDateString().Replace("//", "-") + messages.Headers.DateSent.ToLongTimeString().Replace(":", "-");
                            filename = filename.Replace(",", "-");
                            foreach (var attachment in attachments)
                            {
                                var fullfilename = Path.Combine(sourceBaseFilePath, filename);
                                if (File.Exists(fullfilename))
                                {
                                    fullfilename = "1" + fullfilename;
                                }
                                attachment.Save(new FileInfo(fullfilename + ".pdf"));
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Error For[{index}] -> [{messages?.Headers?.Sender}]  [{e.Message}]");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
                return true;
        }

        private static string parseUsingPDFBox(string filepath)
		{
		    PDDocument doc = null;

            try
            {
                doc = PDDocument.load(filepath);
                PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(doc);
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
		}
	}
}
