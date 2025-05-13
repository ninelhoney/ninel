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

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
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
        

        private void btnDELETE_Click(object sender, EventArgs e)
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
                logs.insertLogs(currentUserName, "Deleted an active student");
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
