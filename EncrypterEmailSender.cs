using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MimeKit;
using MailKit.Net.Smtp;
using OfficeOpenXml;

// Set up log file
using (var logFile = new StreamWriter("log.txt"))
{
    try
    {
        // Read data from Excel file
        using (var package = new ExcelPackage(new FileInfo("data.xlsx")))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                var nationalCode = worksheet.Cells[row, 1].Value.ToString();
                var email = worksheet.Cells[row, 2].Value.ToString();

                // Locate PDF file
                var pdfPath = Directory.GetFiles("pdfs", nationalCode + ".pdf").FirstOrDefault();

                if (pdfPath != null)
                {
                    // Encrypt PDF file
                    var reader = new PdfReader(pdfPath);
                    var outputStream = new MemoryStream();
                    var stamper = new PdfStamper(reader, outputStream);
                    stamper.SetEncryption(PdfWriter.STRENGTH128BITS, nationalCode, nationalCode, PdfWriter.AllowPrinting);
                    stamper.Close();

                    // Send email
                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("Sender Name", "sender@example.com"));
                    message.To.Add(new MailboxAddress("", email));
                    message.Subject = "Encrypted PDF";
                    var attachment = new MimePart("application", "pdf")
                    {
                        Content = new MimeContent(new MemoryStream(outputStream.ToArray()), ContentEncoding.Default),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = nationalCode + ".pdf"
                    };
                    message.Body = new TextPart("plain") { Text = "Please find the encrypted PDF file attached." };
                    message.Attachments.Add(attachment);

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.example.com", 587, false);
                        client.Authenticate("username", "password");

                        try
                        {
                            await client.SendAsync(message);
                            // Log success
                            logFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - Sent encrypted PDF to " + email);
                        }
                        catch (Exception ex)
                        {
                            // Log email sending error
                            logFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - ERROR: Failed to send email to " + email + ": " + ex.ToString());
                        }

                        client.Disconnect(true);
                    }
                }
                else
                {
                    // Log file not found error
                    logFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - ERROR: PDF file not found for national code " + nationalCode);
                }
            }
        }
    }
    catch (Exception ex)
    {
        // Log other errors
        logFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - ERROR: " + ex.ToString());
    }
}
