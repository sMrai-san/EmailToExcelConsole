using Microsoft.Exchange.WebServices.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace EmailToExcel
{
    class Program
    {
        //***********************************
        //HELPER FUNCTIONS
        //***********************************

        //We do not want any html in our email, cleaning it in a function
        public static string HtmlToText(string input)
        {
            input = Regex.Replace(input, "<style>(.|\n)*?</style>", string.Empty);
            input = Regex.Replace(input, @"<xml>(.|\n)*?</xml>", string.Empty);
            input = Regex.Replace(input, @"{[^>]*}", string.Empty);
            input = Regex.Replace(input, @"&nbsp;", string.Empty);
            input = input.Replace("\n", "").Replace("\r", "").Replace("P ", "");
            return Regex.Replace(input, @"<(.|\n)*?>", string.Empty);
        }

        //***********************************
        //MAIN PROGRAM
        //***********************************
        static void Main(string[] args)
        {
            //License (free)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExchangeService _service;

            //How many mails we are reading from newest to last
            int mailsToRead = 100;

            //Needs to be 2 no touchy
            int recordIndex = 2;

            //Defining that we are searching date e.g. 11.11.2021
            string[] format = new string[] { "dd.MM.yyyy" };

            DateTime datetime;
            
            //Connect to Office365
            try
            {
                Console.WriteLine("Registering Exchange connection...");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials("yourOffice365account", "yourpassword")
                };

                //Office365 webservice URL
                _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            }
            catch
            {
                Console.WriteLine("ExchangeService connection failed. Press enter to exit...");
                return;
            }



            //Prepare a class for writing email to the database
            try
            {
                //Create ExcelPackage instance
                ExcelPackage excel = new ExcelPackage();

                //Name the sheet
                var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
                //Headers with colors
                workSheet.Cells[1, 1].Value = "Date Received";
                workSheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                workSheet.Cells[1, 2].Value = "Date TODO";
                workSheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                workSheet.Cells[1, 3].Value = "Sender";
                workSheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                workSheet.Cells[1, 4].Value = "Message";
                workSheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Orange);


                Console.WriteLine("Reading mail...");

                // Read 'n' mails
                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(mailsToRead)))
                {
                    if (email.ConversationTopic != null)
                    {
                        if (email.ConversationTopic.Length > 0)
                        {
			                //if first 10 characters in a message topic includes date we add the message to our workSheet with the automated index
                            if (DateTime.TryParseExact(email.ConversationTopic.Substring(0, 10), format, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out datetime))
                            {
                                workSheet.Cells[recordIndex, 1].Value = email.DateTimeReceived.ToString();
                                workSheet.Cells[recordIndex, 2].Value = email.ConversationTopic.ToString();
                                workSheet.Cells[recordIndex, 3].Value = email.Sender.ToString();
                                var msg = EmailMessage.Bind(_service, email.Id, new PropertySet(EmailMessageSchema.Body));
                                string regexMsg = HtmlToText(msg.Body.Text.ToString());
                                workSheet.Cells[recordIndex, 4].Value = regexMsg;
                                recordIndex++;

                                Console.WriteLine(email.ConversationTopic);
                            }
                        }
                    }
                }

                //Let's autofit the contet to columns
                workSheet.Column(1).AutoFit();
                workSheet.Column(2).AutoFit();
                workSheet.Column(3).AutoFit();
                workSheet.Column(4).AutoFit();

                //file name and path with .xlsx extension 
                string p_strPath = "X:\\Path\\To\\workiswork.xlsx";

                //Delete existing file
                if (File.Exists(p_strPath))
                {
                    File.Delete(p_strPath);
                }

                //Create excel file on physical disk 
                FileStream objFileStrm = File.Create(p_strPath);
                objFileStrm.Close();

                //Write content to excel file 
                File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
                //Close Excel package
                excel.Dispose();

                Console.WriteLine("Excel updated! Press any key to shutdown the program...");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured. \n:" + e.Message);
            }

            Console.ReadLine();
        }


    }
}
