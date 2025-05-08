using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Policy;

namespace ninel
{
    public partial class Form1 : Form
    {
        Form2 form2;

        string[] student = new string[5];
        int i = 0;

        private string currentUserName;

        public Form1(string userName)
        {
            InitializeComponent();
            currentUserName = userName;
        }

        private void btnADD_Click(object sender, EventArgs e)
        {
            string name = txtName.Text.Trim();
            string gender = "";
            string hobbies = "";
            string favcolor = cbFavoriteColor.Text.Trim();
            string address = txtAddress.Text.Trim();
            string email = txtEmail.Text.Trim();
            string birthdate = dtBirthdate.Text.Trim();
            string age = txtAge.Text.Trim();
            string course = cbCourse.Text.Trim();
            string saying = txtSaying.Text.Trim();
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();
            string profilePicture = txtProfilePicture.Text.Trim();

            
            // Validation
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Name cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus(); return;
            }
            else if (int.TryParse(name, out _))
            {
                MessageBox.Show("Name cannot be a number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus(); return;
            }

            gender = radFemale.Checked ? radFemale.Text : radMale.Checked ? radMale.Text : "";
            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cbDancing.Checked) hobbies += cbDancing.Text + ", ";
            if (cbReading.Checked) hobbies += cbReading.Text + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text + ", ";
            hobbies = hobbies.TrimEnd(',', ' ');
            if (string.IsNullOrEmpty(hobbies))
            {
                MessageBox.Show("Please select at least one hobby.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(favcolor))
            {
                MessageBox.Show("Please select a favorite color.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cbFavoriteColor.Focus(); return;
            }

            if (string.IsNullOrEmpty(address))
            {
                MessageBox.Show("Address cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAddress.Focus(); return;
            }

            if (string.IsNullOrEmpty(email) || !email.Contains("@") || !email.Contains("."))
            {
                MessageBox.Show("Please enter a valid email.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtEmail.Focus(); return;
            }

            if (string.IsNullOrEmpty(birthdate))
            {
                MessageBox.Show("Please select a birthdate.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dtBirthdate.Focus(); return;
            }

            if (string.IsNullOrEmpty(age) || !int.TryParse(age, out int ageValue) || ageValue <= 0)
            {
                MessageBox.Show("Please enter a valid age.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtAge.Focus(); return;
            }

            if (string.IsNullOrEmpty(course))
            {
                MessageBox.Show("Please select a course.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cbCourse.Focus(); return;
            }

            if (string.IsNullOrEmpty(saying))
            {
                MessageBox.Show("Saying cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSaying.Focus(); return;
            }

            if (string.IsNullOrEmpty(username))
            {
                MessageBox.Show("Username cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUsername.Focus(); return;
            }

            if (string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Password cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPassword.Focus(); return;
            }

            if (string.IsNullOrEmpty(profilePicture))
            {
                MessageBox.Show("Please browse and select a profile picture.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtProfilePicture.Focus(); return;
            }
            Workbook checkBook = new Workbook();
            checkBook.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet checkSheet = checkBook.Worksheets[0];

            for (int i = 2; i <= checkSheet.LastRow; i++) // skip header
            {
                string existingUsername = checkSheet.Range[i, 11].Value?.Trim();
                string existingPassword = checkSheet.Range[i, 12].Value?.Trim();

                if (string.Equals(existingUsername, username, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(existingPassword, password, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("A user with the same username and password already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Insert to form2 if needed
            Form2 form2 = new Form2(currentUserName);
            form2.insertData(name, gender, hobbies, favcolor, address, email, birthdate, age, course, saying, username, password, "1", profilePicture);

            // Save to Excel
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int r = sh.LastRow + 1;

            sh.Range[r, 1].Value = name;
            sh.Range[r, 2].Value = gender;
            sh.Range[r, 3].Value = hobbies;
            sh.Range[r, 4].Value = address;
            sh.Range[r, 5].Value = favcolor;
            sh.Range[r, 6].Value = email;
            sh.Range[r, 7].Value = birthdate;
            sh.Range[r, 8].Value = age;
            sh.Range[r, 9].Value = course;
            sh.Range[r, 10].Value = saying;
            sh.Range[r, 11].Value = username;
            sh.Range[r, 12].Value = password;
            sh.Range[r, 13].Value = "1"; // active flag
            sh.Range[r, 14].Value = profilePicture; // picture path

            book.SaveToFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("Successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            MyLogs logs = new MyLogs();
            logs.insertLogs(currentUserName, $"Added new user: {name}");

            // Reset form
            btnADD.Visible = true;
            btnUPDATE.Visible = false;

            txtName.Clear(); txtSaying.Clear(); txtAddress.Clear(); txtAge.Clear();
            txtEmail.Clear(); txtUsername.Clear(); txtPassword.Clear(); txtProfilePicture.Clear();

            cbFavoriteColor.SelectedIndex = -1;
            cbCourse.SelectedIndex = -1;
            radMale.Checked = false;
            radFemale.Checked = false;
            cbDancing.Checked = false;
            cbReading.Checked = false;
            cbSinging.Checked = false;

            txtAge.ReadOnly = true;
            txtName.Focus();
        }

        private void btnDISPLAY_Click(object sender, EventArgs e)
        {
            if (form2 == null || form2.IsDisposed)
            {
                Form2 form2 = new Form2(currentUserName);
            }
            form2.Show();
            form2.BringToFront();
        }

        private void btnUPDATE_Click(object sender, EventArgs e)
        {
            
            string name = txtName.Text.Trim();
            string gender = "";
            string hobbies = "";
            string favcolor = cbFavoriteColor.Text.Trim();
            string address = txtAddress.Text.Trim();
            string email = txtEmail.Text.Trim();
            string birthdate = dtBirthdate.Text.Trim();
            string age = txtAge.Text.Trim();
            string course = cbCourse.Text.Trim();
            string saying = txtSaying.Text.Trim();
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();
            string profilePicture = txtProfilePicture.Text.Trim();

            // Validation
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Name cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus(); return;
            }
            else if (int.TryParse(name, out _))
            {
                MessageBox.Show("Name cannot be a number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtName.Focus(); return;
            }

            // Gender selection
            gender = radFemale.Checked ? radFemale.Text : radMale.Checked ? radMale.Text : "";
            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("Please select a gender.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Hobbies selection
            if (cbDancing.Checked) hobbies += cbDancing.Text + ", ";
            if (cbReading.Checked) hobbies += cbReading.Text + ", ";
            if (cbSinging.Checked) hobbies += cbSinging.Text + ", ";
            hobbies = hobbies.TrimEnd(',', ' ');
            if (string.IsNullOrEmpty(hobbies))
            {
                MessageBox.Show("Please select at least one hobby.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Other fields validation
            if (string.IsNullOrEmpty(favcolor)) { MessageBox.Show("Please select a favorite color.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); cbFavoriteColor.Focus(); return; }
            if (string.IsNullOrEmpty(address)) { MessageBox.Show("Address cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtAddress.Focus(); return; }
            if (string.IsNullOrEmpty(email) || !email.Contains("@") || !email.Contains(".")) { MessageBox.Show("Please enter a valid email.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtEmail.Focus(); return; }
            if (string.IsNullOrEmpty(birthdate)) { MessageBox.Show("Please select a birthdate.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); dtBirthdate.Focus(); return; }
            if (string.IsNullOrEmpty(age) || !int.TryParse(age, out int ageValue) || ageValue <= 0) { MessageBox.Show("Please enter a valid age.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtAge.Focus(); return; }
            if (string.IsNullOrEmpty(course)) { MessageBox.Show("Please select a course.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); cbCourse.Focus(); return; }
            if (string.IsNullOrEmpty(saying)) { MessageBox.Show("Saying cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtSaying.Focus(); return; }
            if (string.IsNullOrEmpty(username)) { MessageBox.Show("Username cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtUsername.Focus(); return; }
            if (string.IsNullOrEmpty(password)) { MessageBox.Show("Password cannot be empty.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtPassword.Focus(); return; }
            if (string.IsNullOrEmpty(profilePicture)) { MessageBox.Show("Please browse and select a profile picture.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); txtProfilePicture.Focus(); return; }

            //Workbook checkBook = new Workbook();
            //checkBook.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            //Worksheet checkSheet = checkBook.Worksheets[0];

            //for (int i = 2; i <= checkSheet.LastRow; i++) // skip header
            //{
            //    string existingUsername = checkSheet.Range[i, 11].Value?.Trim();
            //    string existingPassword = checkSheet.Range[i, 12].Value?.Trim();

            //    if (string.Equals(existingUsername, username, StringComparison.OrdinalIgnoreCase) &&
            //        string.Equals(existingPassword, password, StringComparison.OrdinalIgnoreCase))
            //    {
            //        MessageBox.Show("A user with the same username and password already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}

            // Load the Excel file to update
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            

            // Search for the existing row based on username (assuming it's in column 11)
            bool isUpdated = false;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 11].Value.ToString() == username) // Username column
                {
                    // If record is found, update it
                    sh.Range[i, 1].Value = name;
                    sh.Range[i, 2].Value = gender;
                    sh.Range[i, 3].Value = hobbies;
                    sh.Range[i, 4].Value = address;
                    sh.Range[i, 5].Value = favcolor;
                    sh.Range[i, 6].Value = email;
                    sh.Range[i, 7].Value = birthdate;
                    sh.Range[i, 8].Value = age;
                    sh.Range[i, 9].Value = course;
                    sh.Range[i, 10].Value = saying;
                    sh.Range[i, 12].Value = password;
                    sh.Range[i, 13].Value = "1"; // active flag
                    sh.Range[i, 14].Value = profilePicture; // Profile picture path

                    
                    isUpdated = true;

                    
                    break;
                }

                

            }

            // If no existing record was found, add new record
            if (!isUpdated)
            {
                int newRow = sh.LastRow + 1;
                sh.Range[newRow, 1].Value = name;
                sh.Range[newRow, 2].Value = gender;
                sh.Range[newRow, 3].Value = hobbies;
                sh.Range[newRow, 4].Value = address;
                sh.Range[newRow, 5].Value = favcolor;
                sh.Range[newRow, 6].Value = email;
                sh.Range[newRow, 7].Value = birthdate;
                sh.Range[newRow, 8].Value = age;
                sh.Range[newRow, 9].Value = course;
                sh.Range[newRow, 10].Value = saying;
                sh.Range[newRow, 11].Value = username;
                sh.Range[newRow, 12].Value = password;
                sh.Range[newRow, 13].Value = "1"; // active flag
                sh.Range[newRow, 14].Value = profilePicture; // Profile picture path
            }

            // Save changes to Excel
            book.SaveToFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show(isUpdated ? "Successfully updated!" : "Successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            //i added update to my logs
            MyLogs logs = new MyLogs();
            logs.insertLogs(currentUserName, "Updated a student");

            // Reset form
            btnADD.Visible = true;
            btnUPDATE.Visible = false;

            txtName.Clear(); txtSaying.Clear(); txtAddress.Clear(); txtAge.Clear();
            txtEmail.Clear(); txtUsername.Clear(); txtPassword.Clear(); txtProfilePicture.Clear();

            cbFavoriteColor.SelectedIndex = -1;
            cbCourse.SelectedIndex = -1;
            radMale.Checked = false;
            radFemale.Checked = false;
            cbDancing.Checked = false;
            cbReading.Checked = false;
            cbSinging.Checked = false;

            txtAge.ReadOnly = true;
            txtName.Focus();
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            if (d.ShowDialog() == DialogResult.OK)
            {
                txtProfilePicture.Text = d.FileName;

            }

            string profilePath = txtProfilePicture.Text.Trim();
            
        }   

        private void dtBirthdate_ValueChanged(object sender, EventArgs e)
        {
            DateTime birthDate = DateTime.Parse(dtBirthdate.Text);
            int age = DateTime.Now.Year - birthDate.Year;

            if (DateTime.Now < birthDate.AddYears(age))
            {
                age--;
            }

            txtAge.Text = age.ToString();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

}
