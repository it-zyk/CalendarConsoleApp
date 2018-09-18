using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace CalendarConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SendMailByEmail("zyk@cntcc.cn", "测试", "内容"，“”，“”，“exchtrans.cntcc.cn”);
        }


        public static void SendMailByEmail(string toEmail, string mMailTitle, string mMailcontent, string mMailServerUser, string mPassword, string mServerHost)
        {
        
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add(toEmail);
            message.From = new System.Net.Mail.MailAddress(mMailServerUser, "tao");
            message.Subject = mMailTitle;
            message.SubjectEncoding = Encoding.UTF8;
            message.Body = mMailcontent;
            message.BodyEncoding = Encoding.UTF8;
            message.IsBodyHtml = true;
            message.Priority = System.Net.Mail.MailPriority.High;
            try
            {
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient
                {
                    Host = mServerHost,
                    Port = 25,
                    EnableSsl = false,
                    Credentials = new System.Net.NetworkCredential(mMailServerUser, mPassword)
                };
                StringBuilder str = new StringBuilder();
                str.AppendLine("BEGIN:VCALENDAR");
                str.AppendLine("PRODID:-//Schedule a Meeting");
                str.AppendLine("VERSION:2.0");
                str.AppendLine("METHOD:REQUEST");
                str.AppendLine("BEGIN:VEVENT");
                str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", DateTime.Parse("2018-03-23 10:00").ToUniversalTime()));
                str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.Now.ToUniversalTime()));
                str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", DateTime.Parse("2018-03-23 12:00").ToUniversalTime()));
                str.AppendLine("LOCATION: " + "天辰大夏5110会议室");
                str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                str.AppendLine(string.Format("DESCRIPTION:{0}", message.Body));
                str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", message.Body));
                str.AppendLine(string.Format("SUMMARY:{0}", message.Subject));
                str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", message.From.Address));

                str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", message.To[0].DisplayName, message.To[0].Address));

                str.AppendLine("BEGIN:VALARM");
                str.AppendLine("TRIGGER:-PT15M");
                str.AppendLine("ACTION:DISPLAY");
                str.AppendLine("DESCRIPTION:Reminder");
                str.AppendLine("END:VALARM");
                str.AppendLine("END:VEVENT");
                str.AppendLine("END:VCALENDAR");

                System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
                contype.Parameters.Add("method", "REQUEST");
                contype.Parameters.Add("name", "Meeting.ics");
                System.Net.Mail.AlternateView avCal = System.Net.Mail.AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
                message.AlternateViews.Add(avCal);
                client.Send(message);
            }
            catch (System.Net.Mail.SmtpException exception)
            {
                throw new Exception(exception.Message);
            }
        }

    }
}
