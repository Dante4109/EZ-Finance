using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Windows.Forms;

namespace EZ_Finance
{
    class SendMail
    {
        public string fromAddressA { get; set; }
        public string fromAddressB { get; set; }
        public string toAddress { get; set; }

        public string subjectText { get; set; }
        public string bodyText { get; set; }
        public string attachmentPath { get; set; }

        public string usernameMail { get; set; }
        public string passwordMail { get; set; }



        public void sendGmail()
        {

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(fromAddressA, fromAddressB);
            mail.To.Add(toAddress);
            mail.IsBodyHtml = true;
            mail.Subject = subjectText;
            mail.Body = bodyText;
            mail.Priority = MailPriority.High;

            if (attachmentPath != null)
            {
                mail.Attachments.Add(new Attachment(attachmentPath));
            }

            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            //smtp.UseDefaultCredentials = true;
            smtp.Credentials = new System.Net.NetworkCredential(usernameMail, passwordMail);
            smtp.EnableSsl = true;
            //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

            smtp.Send(mail);
            MessageBox.Show("Mail successfully sent!");
        }
    }
}
