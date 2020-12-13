using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {

        //create Contact object
        private Contact co1 = new Contact();

        bool fla = false;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = "";

                using (ContactContext db = new ContactContext())
                {

                    dataGridView1.DataSource = db.Contacts.ToList();

                    //dataGridView1.Columns[0].Visible = false;

                    //dataGridView1.Columns[0].HeaderText = "ID";
                    //dataGridView1.Columns[1].HeaderText = "First Name";
                    //dataGridView1.Columns[2].HeaderText = "Last Name";
                    //dataGridView1.Columns[3].HeaderText = "Email";
                    //dataGridView1.Columns[4].HeaderText = "Phone Number";
                    //dataGridView1.Columns[5].HeaderText = "Birth Date";
                    //dataGridView1.Columns[6].HeaderText = "Address";
                    //dataGridView1.Columns[7].HeaderText = "Description";

                    //foreach (DataGridViewColumn column in dataGridView1.Columns)
                    //{
                    //    column.SortMode = DataGridViewColumnSortMode.Automatic;
                    //}

                    displayGrid();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {

            try
            {
                textBox1.Text = "";

                using (ContactContext db = new ContactContext())
                {

                    dataGridView1.DataSource = db.Contacts.ToList();

                    //dataGridView1.Columns[0].Visible = false;

                    //dataGridView1.Columns[0].HeaderText = "ID";
                    //dataGridView1.Columns[1].HeaderText = "First Name";
                    //dataGridView1.Columns[2].HeaderText = "Last Name";
                    //dataGridView1.Columns[3].HeaderText = "Email";
                    //dataGridView1.Columns[4].HeaderText = "Phone Number";
                    //dataGridView1.Columns[5].HeaderText = "Birth Date";
                    //dataGridView1.Columns[6].HeaderText = "Address";
                    //dataGridView1.Columns[7].HeaderText = "Description";

                    //foreach (DataGridViewColumn column in dataGridView1.Columns)
                    //{
                    //    column.SortMode = DataGridViewColumnSortMode.Automatic;
                    //}

                    displayGrid();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int entryid = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());

                using (ContactContext db = new ContactContext())
                {

                    co1 = db.Contacts.Find(entryid);

                    db.Contacts.Remove(co1);

                    db.SaveChanges();

                    MessageBox.Show("Deleted Successfully");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                Form2 f2 = new Form2();

                f2.textBox3.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                f2.textBox8.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                f2.textBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                f2.textBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                f2.maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();

                if (dataGridView1.CurrentRow.Cells[5].Value.ToString() != "")
                    f2.dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                else
                    f2.dateTimePicker1.CustomFormat = " ";

                f2.textBox6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                f2.textBox7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();

                f2.ShowDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string searchValue = textBox1.Text;

                using (ContactContext db = new ContactContext())
                {

                    var matches = from m in db.Contacts
                                  where
            m.fname.Contains(searchValue) ||
            m.lname.Contains(searchValue) ||
            m.email.Contains(searchValue) ||
            m.mobilephone.Contains(searchValue) ||
            m.birthdate.Contains(searchValue) ||
            m.address.Contains(searchValue) ||
            m.description.Contains(searchValue)
                                  select m;

                    dataGridView1.DataSource = matches.ToList();

                    displayGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void ExportToExcel()
        {

            //Creating an Excel object.
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                //worksheet = workbook.Sheets["Sheet1"];

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                //
                //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    //
                    //worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    worksheet.Cells[1, i] = dataGridView1.Columns[i].HeaderText;
                }

                //Loop through each row and read value from each column.
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //
                    //for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)

                    {
                        //
                        //worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j + 1].Value.ToString();
                    }
                }

                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();

                saveDialog.Filter = "Excel 97-2003 (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx";

                saveDialog.FilterIndex = 1;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);

                excel.Quit();
                GC.Collect();

                Marshal.FinalReleaseComObject(worksheet);

                Marshal.FinalReleaseComObject(workbook);

                Marshal.FinalReleaseComObject(excel);

                //excel.Quit();
                //workbook = null;
                //excel = null;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void displayGrid()
        {
            dataGridView1.Columns[0].Visible = false;

            dataGridView1.Columns[0].HeaderText = "ID";
            dataGridView1.Columns[1].HeaderText = "First Name";
            dataGridView1.Columns[2].HeaderText = "Last Name";
            dataGridView1.Columns[3].HeaderText = "Email";
            dataGridView1.Columns[4].HeaderText = "Phone Number";
            dataGridView1.Columns[5].HeaderText = "Birth Date";
            dataGridView1.Columns[6].HeaderText = "Address";
            dataGridView1.Columns[7].HeaderText = "Description";

            if (dataGridView1.Rows.Count == 0)
            {
                button2.Enabled = false;
                //button3.Enabled = false;
                //textBox1.Enabled = false;
            }
            else
            {
                button2.Enabled = true;
                //button3.Enabled = true;
                //textBox1.Enabled = true;
            }
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            using (ContactContext db = new ContactContext())
            {
                if (fla == false)
                {
                    if (dataGridView1.CurrentCell.ColumnIndex == 1)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.fname).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 2)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.lname).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 3)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.email).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 4)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.mobilephone).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 5)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.birthdate).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 6)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.address).ToList();
                        fla = true;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 7)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderByDescending(s => s.description).ToList();
                        fla = true;
                    }
                }
                else
                {
                    if (dataGridView1.CurrentCell.ColumnIndex == 1)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.fname).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 2)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.lname).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 3)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.email).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 4)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.mobilephone).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 5)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.birthdate).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 6)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.address).ToList();
                        fla = false;
                    }
                    else if (dataGridView1.CurrentCell.ColumnIndex == 7)
                    {
                        dataGridView1.DataSource = db.Contacts.OrderBy(s => s.description).ToList();
                        fla = false;
                    }
                }
            }
        }
    }
}
