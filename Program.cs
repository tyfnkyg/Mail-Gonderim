using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.IO;
using System.Data;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;
using System.Net.Mime;

namespace MailGonderme
{
    class Program
    {
        static void Main(string[] args)
        {
            MailGonder();
        }

        private static void LogGonder(string message)
        {
            Log.LogMessageToFile(message);
        }

        private static void MailGonder()
        {
            try
            {
                SmtpClient smtp = new SmtpClient
                {
                    Credentials = new NetworkCredential(ConfigurationManager.AppSettings["mailSmtpGonderen"], ConfigurationManager.AppSettings["mailSifre"]),
                    EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["mailSmtpSsl"])
                };
                smtp.Port = Convert.ToInt16(ConfigurationManager.AppSettings["mailSmtpPort"]); //465 - smtp.gmail.com  //587 - smtp.live.com-mail.mersus.com
                smtp.Host = ConfigurationManager.AppSettings["mailSmtpHost"];

                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(ConfigurationManager.AppSettings["mailLoginGonderen"], ConfigurationManager.AppSettings["mailLoginGonderenAd"])
                };

                string[] mailCC = ConfigurationManager.AppSettings["mailGonderCC"].Split(','); //CC ile gönderilecek kişiler getiriliyor
                for (int i = 0; i < mailCC.Length; i++)
                {
                    mail.To.Add(mailCC[i]);
                }

                mail.Subject = System.DateTime.Now.ToString("dd.MM.yyy") + " Destek Talepleri Hk.";
                mail.IsBodyHtml = true;

                mail.Body = MailBody();
                mail.Body += GunlukDescSayisi();
                mail.Body += BirOncekiGunKapatilanDesk();
                mail.Body += AylikKapatilanDesc();

                DataTable dtGunlukToplamDestekSayisi = Veri.DtAl(@ConfigurationManager.AppSettings["mailGrafikSorgusu"]);
                ChartCreate(dtGunlukToplamDestekSayisi); //Chart grafik olusturma...

                LinkedResource LinkedImage = new LinkedResource(@"C:\Users\Mersus\Desktop\CoreUygulamalari\Mail gönderme\MailGonderme\myChart.png")
                {
                    ContentId = "MyPic",
                    ContentType = new ContentType(MediaTypeNames.Image.Jpeg)
                };

                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(
                    mail.Body + "<img src=cid:MyPic>",
                  null, "text/html");

                htmlView.LinkedResources.Add(LinkedImage);

                mail.AlternateViews.Add(htmlView);

                smtp.Send(mail);
            }
            catch (Exception e)
            {
                LogGonder("Mail gönderiminde bir hata oluştu " + e.Message);
            }

        }

        public static string GunlukDescSayisi()
        {
            string mailBody;

            mailBody = MailTablo("Şirketler", "mailDestekTalepleri");

            return mailBody;
        }
        public static string BirOncekiGunKapatilanDesk()
        {
            string mailBody;

            mailBody = "<p><u><strong>Bir Önceki Gün Kişi Bazında Kapatılan Destek Sayısı</strong></u></p>";
            mailBody += MailTablo("Kapatanlar", "mailDestekBirOncekiGunKapatılanlar");

            return mailBody;
        }

        public static string AylikKapatilanDesc()
        {
            string mailBody;

            mailBody = "<p><u><strong>Ay Bazında Kapatılan Destek Sayısı</strong></u></p>";
            mailBody += MailTablo("Kapatanlar", "mailDestekAyKapatılanlar");

            return mailBody;
        }

        public static string MailTablo(string isim, string sorguIsmi)
        {
            string mailBody;
            DataTable dtDestekSayisi = Veri.DtAl(@ConfigurationManager.AppSettings[sorguIsmi]);
            mailBody = "<table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 600 + "><tr bgcolor='#4da6ff'><td align='center'><b>" + isim + "</b></td> <td align='center'> <b>Talep Sayısı</b> </td></tr>";

            for (int i = 0; i < dtDestekSayisi.Rows.Count; i++)
            {
                mailBody += "<tr><td align='center'>" + dtDestekSayisi.Rows[i][0] + "</td><td align='center'> " + dtDestekSayisi.Rows[i][1] + "</td> </tr>";
            }

            mailBody += "</table>";
            mailBody += "<br>";

            return mailBody;
        }

        private static void ChartCreate(DataTable dtPerList)
        {
            try
            {
                Chart chart = new Chart
                {
                    DataSource = dtPerList,
                    Width = 500,
                    Height = 450
                };
                chart.Titles.Clear();
                chart.Series.Clear();
                chart.Series.Add("Performans");

                for (int i = 0; i < dtPerList.Rows.Count; i++)
                {
                    if (dtPerList.Rows[i][1].ToString() != "0")
                    {
                        chart.Series[0].Points.AddXY(i + 1, double.Parse(dtPerList.Rows[i][1].ToString()));
                        //chart.Series[0].Points.Add(double.Parse(dtPerList.Rows[i][1].ToString()));
                        chart.Series[0].Points[i].AxisLabel = dtPerList.Rows[i][0].ToString();

                        //chart.Series[0].Color = Color.FromArgb(0, 0, 0);
                        chart.Series[0].BorderColor = Color.FromArgb(164, 164, 164);
                        chart.Series[0].ChartType = SeriesChartType.Column;
                        chart.Series[0].BorderDashStyle = ChartDashStyle.Solid;
                        chart.Series[0].BorderWidth = 1;
                        chart.Series[0].ShadowColor = Color.FromArgb(128, 128, 128);
                        chart.Series[0].ShadowOffset = 1;
                        chart.Series[0].IsValueShownAsLabel = true;

                        chart.Series[0].Font = new Font("Tahoma", 15.0f, FontStyle.Bold);
                        chart.Series[0].BackSecondaryColor = Color.Black;
                        chart.Series[0].LabelForeColor = Color.FromArgb(100, 100, 100);
                    }
                }

                //create chartareas...
                ChartArea ca = new ChartArea
                {
                    Name = "ChartArea",
                    BackColor = Color.White,
                    AxisX = new Axis(),
                    AxisY = new Axis()
                };

                chart.ChartAreas.Add(ca);
                chart.ChartAreas[0].AxisX.LabelStyle.Angle = -70;
                chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Tahoma", 15f);
                chart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.DarkGray;
                chart.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

                //databind...
                chart.DataBind();
                //save result...
                chart.SaveImage(@"C:\Users\Mersus\Desktop\CoreUygulamalari\Mail gönderme\MailGonderme\myChart.png", ChartImageFormat.Png);
                chart.Dispose();
            }
            catch (Exception e)
            {
                LogGonder("Grafik oluşturulurken hata oluştu " + e.Message);
            }
        }

        private static string MailBody()
        {
            string gun = TurkceGun();

            DataTable dtPerList = Veri.DtAl(@ConfigurationManager.AppSettings["mailGonderSorgu"]);
            string body = HtmlTemplate(System.DateTime.Now.ToString("dd.MM.yyyy"), gun, dtPerList.Rows[0][0].ToString());
            return body;
        }

        private static string TurkceGun()
        {
            if (System.DateTime.Now.Day.ToString() == "Monday")
            {
                return "Pazartesi";
            }
            else if (System.DateTime.Now.Day.ToString() == "Tuesday")
            {
                return "Salı";
            }
            else if (System.DateTime.Now.Day.ToString() == "Wednesday")
            {
                return "Çarşamba";
            }
            else if (System.DateTime.Now.Day.ToString() == "Wednesday")
            {
                return "Perşembe";
            }
            else if (System.DateTime.Now.Day.ToString() == "Friday")
            {
                return "Cuma";
            }
            else if (System.DateTime.Now.Day.ToString() == "Saturday")
            {
                return "Cumartesi";
            }
            return "Pazar";
        }
        private static string HtmlTemplate(string date, string gun, string toplam)
        {
            string body = string.Empty;

            using (StreamReader reader = new StreamReader(System.AppDomain.CurrentDomain.BaseDirectory + "EmailTemplate.html"))
            {
                body = reader.ReadToEnd();
            }
            body = body.Replace("{Date}", date);
            body = body.Replace("{Gun}", gun);
            body = body.Replace("{Toplam}", toplam);
            return body;
        }
    }
}
