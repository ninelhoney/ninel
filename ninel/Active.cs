using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ninel
{
    public partial class Active : Form
    {
        private string currentUserName;
        public Active(string userName)
        {
            InitializeComponent();
            currentUserName = userName;

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            bool itemFound = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Skip new row placeholder in DataGridView
                    if (row.IsNewRow) continue;

                    if (row.Cells[0].Value != null &&
                        row.Cells[0].Value.ToString().Equals(txtSearch.Text.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        row.Selected = true;
                        dataGridView1.FirstDisplayedScrollingRowIndex = row.Index; // Scroll to the selected row
                        itemFound = true;
                        break;
                    }
                }

                if (!itemFound)
                {
                    MessageBox.Show("Item was not found in the list. Please try again.", "Search Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during search: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int r = dataGridView1.CurrentCell.RowIndex;

            // Check if Form1 is already open
            Form1 f1 = Application.OpenForms["Form1"] as Form1;

            // Populate Form2 with the data from the selected row in Form1's DataGridView
            if (dataGridView1.Rows[r].Cells[0].Value != null)
                f1.txtName.Text = dataGridView1.Rows[r].Cells[0].Value.ToString();

            // Gender
            string gender = dataGridView1.Rows[r].Cells[1].Value.ToString();
            f1.radMale.Checked = gender == "Male";
            f1.radFemale.Checked = gender == "Female";

            // Hobbies
            string hobbies = dataGridView1.Rows[r].Cells[2].Value.ToString();
            string[] h = hobbies.Split(',');
            f1.cbDancing.Checked = h.Contains("Dancing");
            f1.cbSinging.Checked = h.Contains("Singing");
            f1.cbReading.Checked = h.Contains("Reading");

            // Address
            f1.txtAddress.Text = dataGridView1.Rows[r].Cells[3].Value.ToString();

            // Favorite Color
            f1.cbFavoriteColor.SelectedItem = dataGridView1.Rows[r].Cells[4].Value.ToString();

            // Email
            f1.txtEmail.Text = dataGridView1.Rows[r].Cells[5].Value.ToString();

            // Birthdate
            f1.dtBirthdate.Value = DateTime.TryParse(dataGridView1.Rows[r].Cells[6].Value.ToString(), out DateTime dob) ? dob : DateTime.Now;

            // Age
            f1.txtAge.Text = dataGridView1.Rows[r].Cells[7].Value.ToString();

            // Course
            f1.cbCourse.SelectedItem = dataGridView1.Rows[r].Cells[8].Value.ToString();

            // Saying
            f1.txtSaying.Text = dataGridView1.Rows[r].Cells[9].Value.ToString();

            // Username
            f1.txtUsername.Text = dataGridView1.Rows[r].Cells[10].Value.ToString();

            // Password
            f1.txtPassword.Text = dataGridView1.Rows[r].Cells[11].Value.ToString();

            // Profile Picture
            string picPath = dataGridView1.Rows[r].Cells[13].Value.ToString();
            if (System.IO.File.Exists(picPath))
            {
                f1.txtProfilePicture.Text = picPath;
            }
            else
            {
                f1.txtProfilePicture.Text = "";
            }

            // Hide ADD button, show UPDATE button
            f1.btnADD.Visible = false;
            f1.btnUPDATE.Visible = true;
        }

        private void btnDELETE_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;

                // Update status in DataGridView
                dataGridView1.Rows[selectedIndex].Cells[12].Value = "0";

                // Load the Excel file
                Workbook book = new Workbook();
                book.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
                Worksheet sheet = book.Worksheets[0];


                string username = dataGridView1.Rows[selectedIndex].Cells[10].Value.ToString();

                for (int i = 2; i <= sheet.LastRow; i++)
                {
                    if (sheet.Range[i, 11].Value == username)
                    {
                        sheet.Range[i, 13].Value = "0";
                        break;
                    }
                }
                MyLogs logs = new MyLogs();
                logs.insertLogs(currentUserName, "Deleted an active student");

                // Save changes
                book.SaveToFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");

                MessageBox.Show("Deleted. Status marked as '0'", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnADD_Click(object sender, EventArgs e)
        {
            //add form

            Form1 form1 = new Form1(currentUserName); 
            form1.Show();
        }
    }
}
