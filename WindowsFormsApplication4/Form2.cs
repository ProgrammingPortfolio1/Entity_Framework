using System;
using System.Data.Entity;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApplication4
{
    public partial class Form2 : Form
    {

        //create Contact object
        private Contact co = new Contact();

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //ID
                if (textBox3.Text != "")
                    co.contactID = Convert.ToInt32(textBox3.Text);

                //First Name
                if (textBox8.Text != "")
                    co.fname = textBox8.Text;
                else
                {
                    MessageBox.Show("First Name cannot be left empty");
                    return;
                }

                //Last Name
                if (textBox1.Text != "")
                    co.lname = textBox1.Text;
                else
                {
                    MessageBox.Show("Last Name cannot be left empty");
                    return;
                }


                Regex reg = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
                Match match;

                //Email
                if (textBox2.Text != "")
                {
                    match = reg.Match(textBox2.Text);

                    if (match.Success)
                        co.email = textBox2.Text;
                    else
                    {
                        MessageBox.Show("Email not valid");
                        return;
                    }
                }
                else
                    co.email = textBox2.Text;

                //Phone Number
                co.mobilephone = maskedTextBox2.Text;

                //Birth Date
                if (dateTimePicker1.CustomFormat == " ")
                    co.birthdate = "";
                else if (dateTimePicker1.CustomFormat == "dd/MM/yyyy")
                    co.birthdate = dateTimePicker1.Text;

                //Address
                co.address = textBox6.Text;

                //Description
                co.description = textBox7.Text;

                //create ContactContext object
                using (ContactContext db = new ContactContext())
                {
                    if (textBox3.Text == "") //Create a new database entry
                        db.Contacts.Add(co);
                    else //Update an existing database entry
                        db.Entry(co).State = EntityState.Modified;

                    db.SaveChanges();

                    MessageBox.Show("Database Updated");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Close();

        }

        private void maskedTextBox2_MouseClick(object sender, MouseEventArgs e)
        {
            maskedTextBox2.Select(0, 0);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
                dateTimePicker1.CustomFormat = " ";
        }
    }
}
