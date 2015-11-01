using Microsoft.SharePoint.Client;
using System;
using System.Security;

namespace Sharepoint_List
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool isValid = true;
            textBox1.BackColor = System.Drawing.Color.White;
            textBox2.BackColor = System.Drawing.Color.White;
            textBox3.BackColor = System.Drawing.Color.White;
            textBox4.BackColor = System.Drawing.Color.White;
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
            if (textBox4.Text == "")
            {
                isValid = false;
                textBox4.BackColor = System.Drawing.Color.Red;
            }
            if (isValid)
            {
                try
                {
                    string user = textBox1.Text.ToString();
                    string password = textBox2.Text.ToString();
                    string server = textBox3.Text.ToString();
                    string list = textBox4.Text.ToString();
                    using (ClientContext clientContext = new ClientContext(server))
                    {
                        SecureString passWord = new SecureString();
                        foreach (char c in password.ToCharArray()) passWord.AppendChar(c);                        
                        clientContext.Credentials = new SharePointOnlineCredentials(user , passWord);

                        Microsoft.SharePoint.Client.List oList = clientContext.Web.Lists.GetByTitle(list);
                        Web oWebsite = clientContext.Web;

                        clientContext.Load(oList.Fields);
                        clientContext.ExecuteQuery();

                        richTextBox1.Clear();
                        foreach (Field f in oList.Fields)
                        {                            
                            richTextBox1.AppendText(f.Title + " / " + f.InternalName + " / " + f.Id + " / " + f.Hidden + " / " + f.CanBeDeleted);
                            richTextBox1.AppendText(System.Environment.NewLine);

                        }
                    }
                }
                catch (Exception ex)
                {
                    richTextBox1.Clear();
                    richTextBox1.AppendText("Error: " + ex);
                    richTextBox1.AppendText(System.Environment.NewLine);                    
                }
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            textBox2.Text = string.Empty;
            textBox3.Text = string.Empty;
            textBox4.Text = string.Empty;
            richTextBox1.Clear();
        }
    }
}
