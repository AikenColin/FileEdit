using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OpenPop;
using OpenPop.Pop3;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Common;
using System.Net.Mail;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
            aTimer.Interval = 5000;
            aTimer.Enabled = true;
        }

        private void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        {
            //####################### Get Attachment from Email ###########################################

            ListBox listBox1 = new ListBox();
            var client = new Pop3Client();
            int Port_Number = 995;
            Boolean UseSSL = true;
            string sendBack = string.Empty;
            string returnSubject = string.Empty;

            try
            {
                client.Connect("pop.gmail.com", Port_Number, UseSSL);
                client.Authenticate("caikencode@gmail.com", "XXXXXXX");

                var messageCount = client.GetMessageCount();
                var header = client.GetMessageInfos();

                var Messages = new List<OpenPop.Mime.Message>(messageCount);
                var Headers = new List<MessageHeader>(messageCount);

                for (int i = 0; i < messageCount; i++)
                {
                    OpenPop.Mime.Message getMessage = client.GetMessage(i + 1);
                    Messages.Add(getMessage);
                }

                for (int i = 0; i < messageCount; i++)
                {
                    OpenPop.Mime.Header.MessageHeader getHeader = client.GetMessageHeaders(i + 1);
                    Headers.Add(getHeader);
                }

                foreach (OpenPop.Mime.Message msg in Messages)
                {
                    foreach (var attachment in msg.FindAllAttachments())
                    {
                        string filePath = Path.Combine(@"C:\Users\colin\Desktop\File.txt");

                        if (attachment.FileName.Equals("EditCode.txt"))
                        {
                            FileStream Stream = new FileStream(filePath, FileMode.Create);
                            BinaryWriter BinaryStream = new BinaryWriter(Stream);
                            BinaryStream.Write(attachment.Body);
                            BinaryStream.Close();


                            foreach (var s in Headers)
                            {
                                string Reply = s.ReturnPath.ToString();
                                sendBack = Reply.ToString();

                                string Subby = s.Subject.ToString();
                                returnSubject = Subby.ToString();
                            }

                            if (client.Connected)
                                client.Dispose();
                        }

                        else
                        {
                            if (client.Connected)
                                client.Dispose();

                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("", ex.Message);
            }

            //####################### Edit File ###########################################

            if (System.IO.File.Exists(@"C:\Users\colin\Desktop\File.txt"))
            {
                using (StreamReader sr = new StreamReader(@"C:\Users\colin\Desktop\File.txt"))
                {
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("<div"))
                            {
                                string s = "<div{value here}>";
                                int start = s.IndexOf("<div");
                                int end = s.IndexOf(">");

                                line = line.Remove(start);
                            }

                            if (line.Contains("</div>"))
                            { line = line.Replace("</div>", ""); }

                            if (line.Contains("h1"))
                            { line = line.Replace("h1", "h4"); }

                            if (line.Contains("h2"))
                            { line = line.Replace("h2", "h4"); }

                            if (line.Contains("h3"))
                            { line = line.Replace("h3", "h4"); }

                            if (line.Contains("&#160;"))
                            { line = line.Replace("&#160;", ""); }

                            if (line.Contains("â€™"))
                            { line = line.Replace("â€™", "'"); }

                            if (line.Contains(" â€“"))
                            { line = line.Replace(" â€“", "."); }

                            if (line.Contains("â€"))
                            { line = line.Replace("â€", ""); }

                            if (line.Contains("â€œ"))
                            { line = line.Replace("â€œ", ""); }

                            if (line.Contains("œ"))
                            { line = line.Replace("œ", ""); }

                            if (line.Contains("&#34;"))
                            { line = line.Replace("&#34;", ""); }

                            if (line.Contains("&#38;"))
                            { line = line.Replace("&#38;", ""); }

                            if (line.Contains("&#39;"))
                            { line = line.Replace("&#39;", "'"); }

                            if (line.Contains("class=\"note\""))
                            { line = line.Replace("class=\"note\"", ""); }

                            if (line.Contains("class=\"first\""))
                            { line = line.Replace("class=\"first\"", ""); }

                            if (line.Contains("class=\"last\""))
                            { line = line.Replace("class=\"last\"", ""); }

                            if (line.Contains("class=\"odd\""))
                            { line = line.Replace("class=\"odd\"", ""); }

                            if (line.Contains("class=\"even\""))
                            { line = line.Replace("class=\"even\"", ""); }

                            if (line.Contains("class=\"bot\""))
                            { line = line.Replace("class=\"bot\"", ""); }

                            if (line.Contains("class=\"data-table data-table-simple\""))
                            { line = line.Replace("class=\"data-table data-table-simple\"", "class=\"content-table\""); }

                            if (line.Contains("class=\"data-table data-table-advanced\""))
                            { line = line.Replace("class=\"data-table data-table-advanced\"", "class=\"content-table\""); }

                            if (line.Contains("<tr style=\"height:"))
                            {
                                line = line.Replace("<tr style=\"height: 10px;\">", "<tr>"); line = line.Replace("<tr style=\"height:10px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 20px;\">", "<tr>"); line = line.Replace("<tr style=\"height:20px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 30px;\">", "<tr>"); line = line.Replace("<tr style=\"height:30px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 40px;\">", "<tr>"); line = line.Replace("<tr style=\"height:40px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 50px;\">", "<tr>"); line = line.Replace("<tr style=\"height:50px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 60px;\">", "<tr>"); line = line.Replace("<tr style=\"height:60px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 70px;\">", "<tr>"); line = line.Replace("<tr style=\"height:70px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 80px;\">", "<tr>"); line = line.Replace("<tr style=\"height:80px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 90px;\">", "<tr>"); line = line.Replace("<tr style=\"height:90px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 100px;\">", "<tr>"); line = line.Replace("<tr style=\"height:100px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 110px;\">", "<tr>"); line = line.Replace("<tr style=\"height:110px;\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 120px;\">", "<tr>"); line = line.Replace("<tr style=\"height:120px;\">", "<tr>");

                                line = line.Replace("<tr style=\"height: 10px\">", "<tr>"); line = line.Replace("<tr style=\"height:10px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 20px\">", "<tr>"); line = line.Replace("<tr style=\"height:20px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 30px\">", "<tr>"); line = line.Replace("<tr style=\"height:30px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 40px\">", "<tr>"); line = line.Replace("<tr style=\"height:40px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 50px\">", "<tr>"); line = line.Replace("<tr style=\"height:50px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 60px\">", "<tr>"); line = line.Replace("<tr style=\"height:60px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 70px\">", "<tr>"); line = line.Replace("<tr style=\"height:70px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 80px\">", "<tr>"); line = line.Replace("<tr style=\"height:80px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 90px\">", "<tr>"); line = line.Replace("<tr style=\"height:90px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 100px\">", "<tr>"); line = line.Replace("<tr style=\"height:100px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 110px\">", "<tr>"); line = line.Replace("<tr style=\"height:110px\">", "<tr>");
                                line = line.Replace("<tr style=\"height: 120px\">", "<tr>"); line = line.Replace("<tr style=\"height:120px\">", "<tr>");
                            }

                            if (line.Contains("<th style=\"height:"))
                            {
                                line = line.Replace("<th style=\"height: 10px;\">", "<th>"); line = line.Replace("<th style=\"height:10px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 20px;\">", "<th>"); line = line.Replace("<th style=\"height:20px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 30px;\">", "<th>"); line = line.Replace("<th style=\"height:30px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 40px;\">", "<th>"); line = line.Replace("<th style=\"height:40px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 50px;\">", "<th>"); line = line.Replace("<th style=\"height:50px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 60px;\">", "<th>"); line = line.Replace("<th style=\"height:60px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 70px;\">", "<th>"); line = line.Replace("<th style=\"height:70px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 80px;\">", "<th>"); line = line.Replace("<th style=\"height:80px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 90px;\">", "<th>"); line = line.Replace("<th style=\"height:90px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 100px;\">", "<th>"); line = line.Replace("<th style=\"height:100px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 110px;\">", "<th>"); line = line.Replace("<th style=\"height:110px;\">", "<th>");
                                line = line.Replace("<th style=\"height: 120px;\">", "<th>"); line = line.Replace("<th style=\"height:120px;\">", "<th>");

                                line = line.Replace("<th style=\"height: 10px\">", "<th>"); line = line.Replace("<th style=\"height:10px\">", "<th>");
                                line = line.Replace("<th style=\"height: 20px\">", "<th>"); line = line.Replace("<th style=\"height:20px\">", "<th>");
                                line = line.Replace("<th style=\"height: 30px\">", "<th>"); line = line.Replace("<th style=\"height:30px\">", "<th>");
                                line = line.Replace("<th style=\"height: 40px\">", "<th>"); line = line.Replace("<th style=\"height:40px\">", "<th>");
                                line = line.Replace("<th style=\"height: 50px\">", "<th>"); line = line.Replace("<th style=\"height:50px\">", "<th>");
                                line = line.Replace("<th style=\"height: 60px\">", "<th>"); line = line.Replace("<th style=\"height:60px\">", "<th>");
                                line = line.Replace("<th style=\"height: 70px\">", "<th>"); line = line.Replace("<th style=\"height:70px\">", "<th>");
                                line = line.Replace("<th style=\"height: 80px\">", "<th>"); line = line.Replace("<th style=\"height:80px\">", "<th>");
                                line = line.Replace("<th style=\"height: 90px\">", "<th>"); line = line.Replace("<th style=\"height:90px\">", "<th>");
                                line = line.Replace("<th style=\"height: 100px\">", "<th>"); line = line.Replace("<th style=\"height:100px\">", "<th>");
                                line = line.Replace("<th style=\"height: 110px\">", "<th>"); line = line.Replace("<th style=\"height:110px\">", "<th>");
                                line = line.Replace("<th style=\"height: 120px\">", "<th>"); line = line.Replace("<th style=\"height:120px\">", "<th>");
                            }

                            if (line.Contains("<td style=\"height:"))
                            {
                                line = line.Replace("<td style=\"height: 10px;\">", "<td>"); line = line.Replace("<td style=\"height:10px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 20px;\">", "<td>"); line = line.Replace("<td style=\"height:20px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 30px;\">", "<td>"); line = line.Replace("<td style=\"height:30px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 40px;\">", "<td>"); line = line.Replace("<td style=\"height:40px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 50px;\">", "<td>"); line = line.Replace("<td style=\"height:50px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 60px;\">", "<td>"); line = line.Replace("<td style=\"height:60px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 70px;\">", "<td>"); line = line.Replace("<td style=\"height:70px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 80px;\">", "<td>"); line = line.Replace("<td style=\"height:80px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 90px;\">", "<td>"); line = line.Replace("<td style=\"height:90px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 100px;\">", "<td>"); line = line.Replace("<td style=\"height:100px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 110px;\">", "<td>"); line = line.Replace("<td style=\"height:110px;\">", "<td>");
                                line = line.Replace("<td style=\"height: 120px;\">", "<td>"); line = line.Replace("<td style=\"height:120px;\">", "<td>");

                                line = line.Replace("<td style=\"height: 10px\">", "<td>"); line = line.Replace("<td style=\"height:10px\">", "<td>");
                                line = line.Replace("<td style=\"height: 20px\">", "<td>"); line = line.Replace("<td style=\"height:20px\">", "<td>");
                                line = line.Replace("<td style=\"height: 30px\">", "<td>"); line = line.Replace("<td style=\"height:30px\">", "<td>");
                                line = line.Replace("<td style=\"height: 40px\">", "<td>"); line = line.Replace("<td style=\"height:40px\">", "<td>");
                                line = line.Replace("<td style=\"height: 50px\">", "<td>"); line = line.Replace("<td style=\"height:50px\">", "<td>");
                                line = line.Replace("<td style=\"height: 60px\">", "<td>"); line = line.Replace("<td style=\"height:60px\">", "<td>");
                                line = line.Replace("<td style=\"height: 70px\">", "<td>"); line = line.Replace("<td style=\"height:70px\">", "<td>");
                                line = line.Replace("<td style=\"height: 80px\">", "<td>"); line = line.Replace("<td style=\"height:80px\">", "<td>");
                                line = line.Replace("<td style=\"height: 90px\">", "<td>"); line = line.Replace("<td style=\"height:90px\">", "<td>");
                                line = line.Replace("<td style=\"height: 100px\">", "<td>"); line = line.Replace("<td style=\"height:100px\">", "<td>");
                                line = line.Replace("<td style=\"height: 110px\">", "<td>"); line = line.Replace("<td style=\"height:110px\">", "<td>");
                                line = line.Replace("<td style=\"height: 120px\">", "<td>"); line = line.Replace("<td style=\"height:120px\">", "<td>");
                            }

                            if (line.Contains("style=\"text-align:left") || line.Contains("style=\"text-align: left")
                                || line.Contains("style= \"text-align:left") || line.Contains("style=\" text-align: left"))
                            {
                                line = line.Replace("style=\"text-align:left\"", "");
                                line = line.Replace("style=\"text-align: left\"", "");
                                line = line.Replace("style=\"text-align:left;\"", "");
                                line = line.Replace("style=\"text-align: left;\"", "");
                                line = line.Replace("style=\" text-align:left\"", "");
                                line = line.Replace("style=\" text-align: left\"", "");
                                line = line.Replace("style=\" text-align:left;\"", "");
                                line = line.Replace("style=\" text-align: left;\"", "");
                                line = line.Replace("style= \"text-align:left\"", "");
                                line = line.Replace("style= \"text-align: left\"", "");
                                line = line.Replace("style= \"text-align:left;\"", "");
                                line = line.Replace("style= \"text-align: left;\"", "");
                            }

                            if (line.Contains("<td>") || line.Contains("<td >") || line.Contains("<td  >"))
                            {
                                line = line.Replace("<td>", "<td style=\"text-align:center\">");
                                line = line.Replace("<td >", "<td style=\"text-align:center\">");
                                line = line.Replace("<td  >", "<td style=\"text-align:center\">");
                            }

                            if (line.Contains("<th>") || line.Contains("<th >") || line.Contains("<th  >"))
                            {
                                line = line.Replace("<th>", "<th style=\"text-align:center\">");
                                line = line.Replace("<th >", "<th style=\"text-align:center\">");
                                line = line.Replace("<th  >", "<th style=\"text-align:center\">");
                            }

                            if (line.Contains("<h4") || line.Contains("< h4")
                                && line.Contains("<br") || line.Contains("< br"))
                            {
                                line = line.Replace("<br/>", ""); line = line.Replace("< br/>", "");
                                line = line.Replace("<br/ >", ""); line = line.Replace("< br/ >", "");
                                line = line.Replace("<br />", ""); line = line.Replace("< br />", "");
                                line = line.Replace("<br / >", ""); line = line.Replace("< br / >", "");
                                line = line.Replace("< br/ >", ""); line = line.Replace("< br/  >", "");
                            }

                            if (line.Contains("<table"))
                            {
                                line = line.Replace("<table", "<table border=\"1\"");
                            }

                            if (line.Contains("<ol>"))
                            {
                                line = line.Replace("<ol>", "<ol style=\"list-style-type:lower-alpha\">");
                            }

                            if (line.Contains("<strong>note") || line.Contains("<strong>Note"))
                            {
                                listBox1.Items.Add("<br/><br/>");
                            }

                            if (line.Contains("<strong>important") || line.Contains("<strong>Important"))
                            {
                                listBox1.Items.Add("<br/><br/>");
                            }

                            if (line.Contains("<strong>result") || line.Contains("<strong>Result"))
                            {
                                listBox1.Items.Add("<br/><br/>");
                            }

                            if (line.Contains("<h4>") || line.Contains("<h4 >"))
                            {
                                listBox1.Items.Add("</div>");
                                listBox1.Items.Add("<div class=\"content-box\">");
                                listBox1.Items.Add("<div class=\"information-box\">");
                                listBox1.Items.Add("<div class=\"meatball\"></div>");
                            }

                            listBox1.Items.Add(line);

                            if (line.Contains("</h4>") || line.Contains("</h4 >"))
                            {
                                string p = "</h4>";
                                int start = p.IndexOf("</h4>");

                                listBox1.Items.Add("</div>");
                            }
                        }

                        string look = listBox1.Items[0].ToString();

                        if (look.Contains("</div>") == false)
                        {
                            listBox1.Items.Insert(0, "</div>");
                            listBox1.Items.Insert(0, "<p>SUMMARY</p>");
                            listBox1.Items.Insert(0, "<div class=\"meatball\"></div>");
                            listBox1.Items.Insert(0, "<div class=\"information-box\">");
                            listBox1.Items.Insert(0, "<div class=\"content-box\">");
                            listBox1.Items.Insert(0, "<div id=\"general\">");
                        }

                        if (look.Contains("</div>"))
                        {
                            listBox1.Items.Remove("</div>");
                            listBox1.Items.Insert(0, "<div id=\"general\">");
                        }

                        listBox1.Items.Add("</div> </div>");

                    }
                }
            }
            else
            {
                Application.Restart();
            }

            //####################### Save Attachment ###########################################

            string sPath = @"C:\Users\colin\Desktop\EditFile.txt";

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sPath);
            foreach (var item in listBox1.Items)
            {
                SaveFile.WriteLine(item);
            }

            SaveFile.Close();

            if (System.IO.File.Exists(@"C:\Users\colin\Desktop\File.txt"))
            { System.IO.File.Delete(@"C:\Users\colin\Desktop\File.txt"); }

            else
            {
                Application.Restart();
            }

            //####################### Send Back Edited File ###########################################

            try
            {

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress("caikencode@gmail.com");
                mail.To.Add(sendBack);
                mail.Subject = returnSubject;
                mail.Body = "Have a Nice Day!!";

                System.Net.Mail.Attachment newAttachment;
                newAttachment = new System.Net.Mail.Attachment(@"C:\Users\colin\Desktop\EditFile.txt");
                mail.Attachments.Add(newAttachment);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("caikencode@gmail.com", "caiken121");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
