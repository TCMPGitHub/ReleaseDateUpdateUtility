using System;
using System.ComponentModel;
using System.IO;
using System.Net.Mail;
using System.Configuration;

namespace ReleaseDateUpdateUtility
{
    class Program
    {
        static bool mailSent = false;
        static string mailbody = string.Empty;
        static string DBBassConn = string.Empty;
        static string FilePath = string.Empty;
        static void Main(string[] args)
        {
            ReleaseDateUpdate lib = new ReleaseDateUpdate();
            lib.DBBassConn = ConfigurationManager.ConnectionStrings["BassWeb"].ConnectionString;
            string path = ConfigurationManager.AppSettings["FilePath"];
            string filename = Path.Combine(path, "ReleaseDateChanges.csv");
            if (File.Exists(filename))
            {
                lib.FilePath = filename;
                mailbody = DateTime.Now.ToString() + " Start Releasedate update process..." + Environment.NewLine;
                string retmessge = lib.Import();
                mailbody = mailbody + DateTime.Now.ToString()  + " " + retmessge + Environment.NewLine;
                //rename file with timestamp
                var newFile = Path.Combine(path, "ReleaseDateChanges" + DateTime.Now.ToString("yyyyMMddhhss") + ".csv");
                System.IO.File.Move(filename, newFile);
                string rename = " ReleaseDateChanges.csv rename to ReleaseDateChanges" + DateTime.Now.ToString("yyyyMMddhhss") + ".csv";
                LogWriter.LogMessageToFile(rename);
                mailbody = mailbody + DateTime.Now.ToString() + rename + Environment.NewLine;
            }
            else
            {
                mailbody = mailbody + DateTime.Now.ToString() + " " + filename + " File Not Found..." + Environment.NewLine;
            }

            LogWriter.LogMessageToFile("Send email to TCMP ..." + DateTime.Now.ToString());
            mailbody = mailbody + DateTime.Now.ToString() + " Send email to TCMP ...";
            SendMail();
            return;
        }

        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                LogWriter.LogMessageToFile(" Message sent.");
            }
            mailSent = true;
        }

        private static void SendMail()
        {
            SmtpClient client = new SmtpClient("smtp.cdcr.ca.gov");
            MailAddress from = new MailAddress("FDCDBA3_TCMP@cdcr.ca.gov", "TCMP", System.Text.Encoding.UTF8);
            // Set destinations for the e-mail message.
            MailAddressCollection toAddresses = new MailAddressCollection();
            toAddresses.Add(new MailAddress("Quentin.Miller@cdcr.ca.gov"));
            //toAddresses.Add(new MailAddress("carol.xu@cdcr.ca.gov"));

            // Specify the message content.
            MailMessage message = new MailMessage();
            message.From = from;
            message.To.Add(toAddresses.ToString());
            message.CC.Add("carol.xu@cdcr.ca.gov");
            message.Body = "Release Data updated Completed at " + DateTime.Now + ":";


            // Include some non-ASCII characters in body and subject.
            message.Body = mailbody + Environment.NewLine;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.Subject = "Release Date Changed!";
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            // Set the method that is called back when the send operation ends.
            client.SendCompleted += new 
            SendCompletedEventHandler(SendCompletedCallback);
            // The userState can be any object that allows your callback 
            // method to identify this send operation.
            string userState = "Update Data";
            client.SendAsync(message, userState);
            string answer = Console.ReadLine();
            // Clean up.
            message.Dispose();
            System.Environment.Exit(0);
            //return;
        }
    }
}
