using Emailer.Properties;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using Tortuga.Chain;

namespace Emailer
{
    class Program
    {
        static void Main(string[] args)
        {
            var settings = Settings.Default;

            var ds = new AccessDataSource(settings.DatabaseConnectionString);
            ds.TestConnection();

            //Connection strings:
            //https://www.connectionstrings.com/access/

            //If you see The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine
            //https://www.connectionstrings.com/the-microsoft-ace-oledb-12-0-provider-is-not-registered-on-the-local-machine/
            //This didn't work for me. I can only open MDB files, not accdb files.
            //Fortunately it is easy to change the type using Save As in Access.

            var filter = new { EmailSent = false }; //This could also be a SQL WHERE string
            var newRows = ds.From<Table1>(filter).ToCollection().Execute();

            if (newRows.Count > 0)
            {
                using (var emailClient = new SmtpClient())
                {
                    emailClient.Connect(settings.SmtpServer, settings.SmtpPort, SecureSocketOptions.Auto);

                    //Remove any OAuth functionality as we won't be using it.
                    emailClient.AuthenticationMechanisms.Remove("XOAUTH2");

                    emailClient.Authenticate(settings.SmtpUsername, settings.SmtpPassword);

                    foreach (var row in newRows)
                    {
                        //Create and send the message
                        var message = new MimeMessage();
                        message.To.Add(new MailboxAddress(settings.ToAddress));
                        message.From.Add(new MailboxAddress(settings.FromAddress));

                        message.Subject = "This is a test for record " + row.Id;

                        //We will say we are sending HTML. But there are options for plaintext etc.
                        message.Body = new TextPart(TextFormat.Html)
                        {
                            Text = "Hello " + row.Name
                        };
                        emailClient.Send(message);
                        Console.WriteLine("Send email for " + row.Id);

                        //Update the database
                        row.EmailSent = true;
                        ds.Update(row).Execute();
                    }

                    emailClient.Disconnect(true);
                }
            }
            else
            {
                Console.WriteLine("No new rows");
            }
        }
    }

    [Table("Table1")] //The table name
    class Table1
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public bool EmailSent { get; set; }
    }
}
