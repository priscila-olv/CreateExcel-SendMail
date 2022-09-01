using ClosedXML.Excel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Text;

namespace CreateExcel_SendMail
{
    public class Report
    {
        private System.Data.DataTable Result { get; set; }
        public StringBuilder Log { get; set; }
        public Report()
        {
            Log = new StringBuilder();
        }

        public void Generate()
        {
            GetData();
            PrepareFile();
            SendMail();
            DeleteOldFiles();
        }

        private void GetData()
        {
            Result = new System.Data.DataTable();

            var datasource = @"(LocalDb)\MSSQLLocalDB";
            var database = "DbItems";
            var username = "sa";
            var password = "123";

            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;

            using (SqlConnection con = new SqlConnection(connString))
            {
                con.Open();
                using (SqlCommand command = new SqlCommand(@"select * from Product", con))
                {
                    SqlDataAdapter da = new SqlDataAdapter(command);

                    da.Fill(Result);
                    con.Close();
                    da.Dispose();
                }
            }
        }
        private void PrepareFile()
        {
            var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Data");
            var table = sheet.Cell(1, 1).InsertTable(Result, "Data", true);
            sheet.Columns(1, 10).AdjustToContents();
            var ptSheet = workbook.Worksheets.Add("Table");
            ptSheet.Columns(1, 10).AdjustToContents();
            var pt = ptSheet.PivotTables.Add("Table", ptSheet.Cell(1, 1), table.AsRange());
            pt.RowLabels.Add("perfil");
            pt.RowLabels.Add("item");
            pt.RowLabels.Add("status perfil x item");
            workbook.SaveAs($"Status Product {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx");
        }

        private void SendMail()
        {
            string fileName = $"Status Product {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx";
            MailMessage mailMessage = new MailMessage();

            mailMessage.From = new MailAddress("email@hotmail.com");
            mailMessage.To.Add(new MailAddress("email2@hotmail.com"));

            //Copy:
            //mailMessage.CC.Add(new System.Net.Mail.MailAddress("copy@email.com"));

            //Hidden Copy:
            //mailMessage.Bcc.Add(new System.Net.Mail.MailAddress("hidden.copy@email.com"));

            mailMessage.Subject = "Subject";
            mailMessage.Body = "Body";
            mailMessage.IsBodyHtml = false;
            mailMessage.Attachments.Add(new Attachment(fileName));

            using (var smtp = new SmtpClient())
            {
                smtp.Host = "smtp-mail.outlook.com";
                smtp.Port = 25;
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential("email@hotmail.com", "password");
                smtp.Send(mailMessage);
            }
        }
        private void DeleteOldFiles()
        {
            string arquivosExcluidos = "";
            string[] a = Directory.GetFiles(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "*.xlsx");
            foreach (string arquivo in a)
            {
                try
                {
                    FileInfo infoArquivo = new FileInfo(arquivo);
                    if (infoArquivo.CreationTime < DateTime.Now.AddDays(-30))
                    {
                        File.Delete(arquivo);
                        arquivosExcluidos += arquivo;
                    }
                }
                catch { }
            }
            if (!string.IsNullOrEmpty(arquivosExcluidos))
                Log.AppendLine(arquivosExcluidos);
        }
    }
}
