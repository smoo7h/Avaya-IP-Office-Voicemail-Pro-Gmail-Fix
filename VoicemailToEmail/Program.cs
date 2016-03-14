using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Threading;
namespace VoicemailToEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> fileEnt = Directory.GetDirectories(@"C:\Program Files (x86)\Avaya\IP Office\Voicemail Pro\VM\Accounts").ToList();

            foreach (var dir in fileEnt)
            {
                FileSystemWatcher watcher = new FileSystemWatcher(dir);

                watcher.Created += new FileSystemEventHandler(OnChange);

                watcher.EnableRaisingEvents = true;

            }
            Console.WriteLine("Voicemail watcher running! Input 'q' to end watch");
            while (Console.Read() != 'q') ;
        }
        private static void OnChange(object source, FileSystemEventArgs e)
        {
            try
            {
                string filePath = e.FullPath;
                List<string> parentDirs = filePath.Split('\\').ToList();
                string userName = parentDirs[7].ToString();
                DirectoryInfo info = new DirectoryInfo(@"C:\Program Files (x86)\Avaya\IP Office\Voicemail Pro\VM\Accounts\" + userName);
                userName = userName.Replace(" ", ".");
                //Error occured System.FormatException : An invalid character was found in the mail header: ' '. Line 85 where trying to attach the userName
                Thread.Sleep(1000);


              
                if (!userName.Equals(null))
                {
                    userName += "@yourdomain.org";//put your domain name here 
                }

                List<FileInfo> files = info.GetFiles().OrderByDescending(p => p.CreationTime).ToList();
                string wavPath = files[0].DirectoryName + "\\" + files[0].ToString();
                string date = DateTime.Now.ToString();


                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                System.Net.Mail.Attachment mail = new System.Net.Mail.Attachment(wavPath);

                message.Subject = "A new voicemail has been recieved at: " + date;
                message.Attachments.Add(mail);
                message.To.Add(userName);
                message.From = new System.Net.Mail.MailAddress("voicemail@greensaver.org");
                message.Body = "You have recieved a new voicemail. Please see the attached document and have a great day!";

                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com");

                smtp.Port = 25;
                smtp.Credentials = new System.Net.NetworkCredential("", "");//add your email and password 
                smtp.EnableSsl = true;
                smtp.Send(message);

                Console.WriteLine("An email has been sent to " + userName);
                message.Dispose();
                mail.Dispose();
                smtp.Dispose();
            }
            catch
            {
                using (StreamWriter writer = new StreamWriter(@"C:\EmailSentErrors\ErrorDoc.txt", true))
                {
                    writer.WriteLine("Error sending email on " + DateTime.Now.ToString());
                    writer.WriteLine("------------------------------------------");
                }
            }
        }
    }
}
