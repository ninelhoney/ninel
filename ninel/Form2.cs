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
    public partial class Form2 : Form

    {
        
        private string currentUserName;
        Logs logs = new Logs();
        public Form2(string userName)
        {
            InitializeComponent();
            LoadExcelFile();
            currentUserName = userName;
        }
        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dataGridView1.DataSource = dt;
        }

        public void insertData(string name, string gender, string hobbies, string favColor,
                       string address, string email, string birthdate, string age,
                       string course, string saying, string username, string password,
                       string status, string profilePicture)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            DataRow newRow = dt.NewRow();

            newRow[0] = name;
            newRow[1] = gender;
            newRow[2] = hobbies;
            newRow[3] = favColor;
            newRow[4] = address;
            newRow[5] = email;
            newRow[6] = birthdate;
            newRow[7] = age;
            newRow[8] = course;
            newRow[9] = saying;
            newRow[10] = username;
            newRow[11] = password;
            newRow[12] = status;
            newRow[13] = profilePicture;

            dt.Rows.Add(newRow);
        }
        public void update(string name, string gender, string hobbies, string favColor,
                      string address, string email, string birthdate, string age,
                      string course, string saying, string username, string password,
                      string status, string profilePicture)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            DataRow newRow = dt.NewRow();

            newRow[0] = name;
            newRow[1] = gender;
            newRow[2] = hobbies;
            newRow[3] = favColor;
            newRow[4] = address;
            newRow[5] = email;
            newRow[6] = birthdate;
            newRow[7] = age;
            newRow[8] = course;
            newRow[9] = saying;
            newRow[10] = username;
            newRow[11] = password;
            newRow[12] = status;
            newRow[13] = profilePicture;

            dt.Rows.Add(newRow);
        }
        private void btnDELETE_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;

                // Update status in DataGridView
                dataGridView1.Rows[selectedIndex].Cells[12].Value = "0";

                // Load the Excel file
                Workbook book = new Workbook();
                book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
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
               
                // Save changes
                book.SaveToFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");

                MessageBox.Show("Deleted. Status marked as '0'", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                MyLogs logs = new MyLogs();
                logs.insertLogs(currentUserName, "Deleted a student");
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView1_CellMouseDoubleClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            int r = dataGridView1.CurrentCell?.RowIndex ?? -1;
            if (r < 0 || r >= dataGridView1.Rows.Count) return;

            Form1 f1 = Application.OpenForms["Form1"] as Form1;

            if (f1 == null)
            {
                if (string.IsNullOrEmpty(currentUserName)) return;
                f1 = new Form1(currentUserName);
                f1.Show();
            }

            // Check each cell and assign safely
            var cells = dataGridView1.Rows[r].Cells;

            f1.txtName.Text = cells[0]?.Value?.ToString() ?? "";

            string gender = cells[1]?.Value?.ToString() ?? "";
            f1.radMale.Checked = gender == "Male";
            f1.radFemale.Checked = gender == "Female";

            string hobbies = cells[2]?.Value?.ToString() ?? "";
            string[] h = hobbies.Split(',');
            f1.cbDancing.Checked = h.Contains("Dancing");
            f1.cbSinging.Checked = h.Contains("Singing");
            f1.cbReading.Checked = h.Contains("Reading");

            f1.txtAddress.Text = cells[3]?.Value?.ToString() ?? "";
            f1.cbFavoriteColor.SelectedItem = cells[4]?.Value?.ToString() ?? "";
            f1.txtEmail.Text = cells[5]?.Value?.ToString() ?? "";

            string birthdate = cells[6]?.Value?.ToString();
            f1.dtBirthdate.Value = DateTime.TryParse(birthdate, out DateTime dob) ? dob : DateTime.Now;

            f1.txtAge.Text = cells[7]?.Value?.ToString() ?? "";
            f1.cbCourse.SelectedItem = cells[8]?.Value?.ToString() ?? "";
            f1.txtSaying.Text = cells[9]?.Value?.ToString() ?? "";
            f1.txtUsername.Text = cells[10]?.Value?.ToString() ?? "";
            f1.txtPassword.Text = cells[11]?.Value?.ToString() ?? "";

            string picPath = cells[13]?.Value?.ToString() ?? "";
            f1.txtProfilePicture.Text = System.IO.File.Exists(picPath) ? picPath : "";

            f1.btnADD.Visible = false;
            f1.btnUPDATE.Visible = true;
        }
        

        private void btnSearch_Click_1(object sender, EventArgs e)
        {

            dataGridView1.ClearSelection();
            bool itemFound = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Skip new row placeholder
                    if (row.IsNewRow) continue;

                    if (row.Cells[0].Value != null &&
                        row.Cells[0].Value.ToString().IndexOf(txtSearch.Text.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        row.Selected = true;
                        // Optional: Scroll to the first match only
                        if (!itemFound)
                        {
                            dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
                        }
                        itemFound = true;
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

        private void btnCLOSE_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
