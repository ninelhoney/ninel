using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ninel
{
    public partial class Login : Form

    {   
        MyLogs logs = new MyLogs();
       
        public Login()
        {
            InitializeComponent();    
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                txtPassword.UseSystemPasswordChar = true; // Show password
            }
            else
            {
                txtPassword.UseSystemPasswordChar = false; // Hide password
            }
        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            //load excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            bool loginSuccess = false;

            for (int i = 2; i <= sheet.LastRow; i++)
            {
                string storedUsername = sheet.Range[i, 11].Value?.Trim(); 
                string storedPassword = sheet.Range[i, 12].Value?.Trim(); 
                string accountStatus = sheet.Range[i, 13].Value?.Trim();  

                if (storedUsername == txtUsername.Text.Trim() && storedPassword == txtPassword.Text.Trim())
                {
                    //validate if account is inactive
                    if (accountStatus == "0")
                    {
                        MessageBox.Show("Your account is inactive. Login Failed", "Account Inactive", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        loginSuccess = true;
                        txtUsername.Clear(); txtPassword.Clear();
                        break;
                    }

                    string profilePath = sheet.Range[i, 14].Text;
                    string name = storedUsername;

                    MessageBox.Show("Login successful", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Hide();

                    logs.insertLogs(storedUsername, "Successfully logged in!");

                    Dashboard dashboard = new Dashboard(name, profilePath);
                    dashboard.ShowDialog();
                    loginSuccess = true;
                    this.Close();
                    break;
                }

            }

            if (!loginSuccess)
            {
                // Validate username and password fields
                if (string.IsNullOrWhiteSpace(txtPassword.Text) && string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    MessageBox.Show("Username and password cannot be empty.","Login Failed",  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    MessageBox.Show("Username cannot be empty.","Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (string.IsNullOrWhiteSpace(txtPassword.Text))
                {
                    MessageBox.Show("Password cannot be empty.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //close
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}