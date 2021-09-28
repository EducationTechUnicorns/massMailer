using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace sendMails
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public DataTable res = new DataTable();

        private void button1_Click(object sender, EventArgs e)
        {

            DataSet dset = new DataSet();

            string filepath = "c://Processo-Email-CSV.csv";
            res = ConvertCSVtoDataTable(filepath);
            dset.Tables.Add(res);
            dataGridView1.DataSource = res;

        }
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            string[] headers = sr.ReadLine().Split(',');
            DataTable dt = new DataTable();
            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        public DataTable getLinq(DataTable dt1, DataTable dt2)
        {
            DataTable dtMerged =
                 (from a in dt1.AsEnumerable()
                  join b in dt2.AsEnumerable()
                           on
                             a["processo"].ToString() equals b["processo"].ToString()
                         into g
                  where g.Count() > 0
                  select a).CopyToDataTable();
            return dtMerged;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string to = "";         
            int count = 0;
            foreach (DataRow row in res.Rows)
            {
                to = row["MailEE"].ToString();

                var smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential("antoniocsantos@agrupajunqueira.pt", "m1$h4n4t4$h4"),
                    EnableSsl = true,
                };
                string body = "Caro Encarredado de Educação \r\n" + "As senhas de acesso do educando à plataforma google da escola são as seguintes \r\n"
                    + "Utilizador: " + row["User"] + "\r\n" + "password: " + row["Pass"] + "\r\n" + "Obrigado";

                MessageBox.Show(to);
                smtpClient.Send("geral@agrupajunqueira.pt", to, "Acesso Google Escola", body);
                count++;
                label1.Text = count.ToString();
            }
        }
    }
}
