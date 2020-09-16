using CT;
using CT.Collections.Generic;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;

namespace AverageSpreadsExcelReport
{
    /// <summary>
    /// Provides (html) email sending
    /// </summary>
    public class Mailer
    {
        #region private implementation

        string Username;
        string Password;
        int Port;
        string Server;
        public readonly MailAddress From;
        public readonly MailAddress[] To;
        public readonly MailAddress[] Bcc;
        
        /// <summary>Loads the addresses from a address array. Each address is checked for validity and uniqueness.</summary>
        /// <param name="addresses">The addresses.</param>
        /// <returns>Returns an array of validated unique MailAddresses</returns>
        MailAddress[] LoadAddresses(string[] addresses)
        {
            Set<MailAddress> receipients = new Set<MailAddress>();
            foreach (string address in addresses)
            {
                receipients.Add(new MailAddress(address));
            }
            return receipients.ToArray();
        }
        #endregion

        #region constructor        
        /// <summary>Initializes a new instance of the <see cref="Mailer"/> class using the given configuration.</summary>
        /// <param name="config">The configuration.</param>
        public Mailer(Ini config)
        {
            if (!config.GetValue("Mail", "Server", ref Server) ||
                !config.GetValue("Mail", "Port", ref Port) ||
                !config.GetValue("Mail", "Password", ref Password) ||
                !config.GetValue("Mail", "Username", ref Username))
            {
                throw new Exception("[Mail] configuration is invalid!");
            }
            //TODO: Optional Display Name for Sender
            string from = config.ReadSetting("Mail", "From");
            From = new MailAddress(from, from);
            To = LoadAddresses(config.ReadSection("SendTo", true));
            Bcc = LoadAddresses(config.ReadSection("BlindCarbonCopy", true));

            if (To.Length == 0) throw new Exception("No recepient (SendTo) address.");
        }

        #endregion

        #region public functionality
        /// <summary>Sends an email.</summary>
        /// <param name="subject">The subject.</param>
        /// <param name="bodyHtml">The body HTML.</param>
        /// <param name="bodyText">The body text.</param>
        /// <param name="files">Attached files</param>
        public void SendMail(string subject, string bodyHtml, string bodyText, List<string> files = null)
        {
            using (MailMessage message = new MailMessage())
            {
                foreach (MailAddress a in To) message.To.Add(a);
                foreach (MailAddress a in Bcc) message.Bcc.Add(a);
                message.Subject = subject;
                message.Body = bodyHtml;
                message.IsBodyHtml = true;
                message.From = From;

                // Attach files
                if (files != null)
                {
                    foreach (string fileName in files)
                    {
                        message.Attachments.Add(new Attachment(fileName));
                    }
                }

                //TODO: Check the alternate view setting
                //message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(bodyText));

                using (SmtpClient client = new SmtpClient(Server, Port))
                {
                    client.UseDefaultCredentials = true;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(Username, Password);
                    client.Send(message);
                }
            }
        }
        #endregion
    }
}