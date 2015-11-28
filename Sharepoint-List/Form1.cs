using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Security;
using System.Threading;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.Xml;
using SP = Microsoft.SharePoint.Client;
using System.Drawing;

namespace Sharepoint_List
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        DataTable dt = new DataTable();
        List<string> splist = new List<string>();
        public Form1()
        {
            InitializeComponent();
            textBox2.KeyDown += new KeyEventHandler(textBox2_KeyDown);
            // Create the events for the Background Worker.
            if (worker.IsBusy != true)
            {
                worker.WorkerReportsProgress = true;
                worker.WorkerSupportsCancellation = true;
                worker.DoWork += new DoWorkEventHandler(worker_DoWork);
                worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
                worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dt.Columns.Add("Title");
            dt.Columns.Add("Internal Name");
            dt.Columns.Add("ID");
            dt.Columns.Add("Is Hidden");
            dt.Columns.Add("Can Be Deleted");
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            comboBox1.BackColor = Color.Red;
        }

        //Background Worker Do Work
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string user = textBox1.Text.ToString().Trim();
                string password = textBox2.Text.ToString().Trim();
                string server = textBox3.Text.ToString().TrimEnd('/');
                using (ClientContext clientContext = new ClientContext(server))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(user, passWord);

                    Microsoft.SharePoint.Client.List oList = clientContext.Web.Lists.GetByTitle(splist[0]);
                    Web oWebsite = clientContext.Web;

                    clientContext.Load(oList.Fields);
                    clientContext.ExecuteQuery();

                    Invoke(new Action(() => label5.Text = ""));
                    foreach (Field f in oList.Fields)
                    {
                        dt.Rows.Add(f.Title, f.InternalName, f.Id, f.Hidden, f.CanBeDeleted);
                    }
                    for (int i = 1; i <= dt.Rows.Count; i++)
                    {
                        if ((worker.CancellationPending == true))
                        {
                            e.Cancel = true;
                            break;
                        }
                        else
                            worker.ReportProgress(Convert.ToInt32(i * 100 / dt.Rows.Count));
                    }
                }
            }
            catch (Exception ex)
            {
                Invoke(new Action(() => label5.Text = ""));
                Invoke(new Action(() => label5.Text = "Error: " + ex));
            }
        }

        //Background Worker Progress Changed
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label6.Text = e.ProgressPercentage.ToString() + "%";
        }

        //Background Worker Completed
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();

            if ((e.Cancelled == true))
            {
                label5.Text = "Search Canceled By User!";
            }

            else if (!(e.Error == null))
            {
                label5.Text = "Error: " + e.Error.Message;
            }
            else
            {
                this.dataGridView1.DataSource = dt;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        //Search Button
        private void button1_Click(object sender, EventArgs e)
        {
            //Start Worker
            worker.RunWorkerAsync();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button4_Click(this, new EventArgs());
            }
        }

        //Clear Button
        private void button2_Click(object sender, EventArgs e)
        {            
            progressBar1.Value = 0;
            label6.Text = "%";
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            dt.Rows.Clear();
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        //Cancel Button
        private void button3_Click(object sender, EventArgs e)
        {
            if (worker.WorkerSupportsCancellation == true)
            {
                worker.CancelAsync();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label5_Click_1(object sender, EventArgs e)
        {

        }

        //Login Buton
        private void button4_Click(object sender, EventArgs e)
        {
            bool isValid = true;
            textBox1.BackColor = System.Drawing.Color.White;
            textBox2.BackColor = System.Drawing.Color.White;
            textBox3.BackColor = System.Drawing.Color.White;
            if (textBox1.Text == "")
            {
                isValid = false;
                textBox1.BackColor = System.Drawing.Color.Red;
            }
            if (textBox2.Text == "")
            {
                isValid = false;
                textBox2.BackColor = System.Drawing.Color.Red;
            }
            if (textBox3.Text == "")
            {
                isValid = false;
                textBox3.BackColor = System.Drawing.Color.Red;
            }
            if (isValid)
            {
                string user = textBox1.Text.ToString().Trim();
                string password = textBox2.Text.ToString().Trim();
                string server = textBox3.Text.ToString().TrimEnd('/');
                using (ClientContext clientContext = new ClientContext(server))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(user, passWord);
                    Web oWebsite = clientContext.Web;
                    ListCollection collList = oWebsite.Lists;
                    IEnumerable<SP.List> resultCollection = clientContext.LoadQuery(
                    collList.Include(
                        list => list.Title,
                        list => list.Id));
                    clientContext.ExecuteQuery();
                    foreach (SP.List oList in resultCollection)
                    {
                        comboBox1.Items.Add(oList.Title);
                    }
                    comboBox1.BackColor = Color.LightGreen;
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt.Rows.Clear();
            splist.Clear();
            splist.Add(comboBox1.SelectedItem.ToString());
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
        }
    }
}
