using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ninel
{
    public partial class Dashboard : Form
    {
        private string currentUserName;

        public Dashboard(string name, string path)
        {
            InitializeComponent();
            currentUserName = name;


            // Load the Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            // Retrieve the username 
            lblName.Text = "Welcome, " + name;

            try
            {
                pictureBox1.Image = Image.FromFile(path);
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
                MessageBox.Show("Error loading profile picture:\n" + ex.Message, "Image Error");
            }

            // Count active and inactive students
            int activeStudentCount = 0;
            int inactiveStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string status = sh.Range[i, 13].Value?.ToString().Trim();
                if (status == "1") activeStudentCount++;
                else if (status == "0") inactiveStudentCount++;
            }
            lblActive.Text = activeStudentCount.ToString();
            lblInactive.Text = inactiveStudentCount.ToString();

            // Count the male students
            int maleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }
            }

            // Count the female students
            int femaleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }
            }

            // Count hobbies
            int dancingHobbiesCount = 0;
            int singingHobbiesCount = 0;
            int readingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string hobby = sh.Range[i, 3].Value.ToString();
                if (hobby == "Dancing") dancingHobbiesCount++;
                if (hobby == "Singing") singingHobbiesCount++;
                if (hobby == "Reading") readingHobbiesCount++;
            }
            lblDancing.Text = dancingHobbiesCount.ToString();
            lblSinging.Text = singingHobbiesCount.ToString();
            lblReading.Text = readingHobbiesCount.ToString();

            // Count favorite colors
            int pinkColorCount = 0;
            int blackColorCount = 0;
            int whiteColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string color = sh.Range[i, 5].Value.ToString();
                if (color == "Pink") pinkColorCount++;
                if (color == "Black") blackColorCount++;
                if (color == "White") whiteColorCount++;
            }
            lblPink.Text = pinkColorCount.ToString();
            lblBlack.Text = blackColorCount.ToString();
            lblWhite.Text = whiteColorCount.ToString();

            // Count courses
            int bsitCourseCount = 0;
            int bsedCourseCount = 0;
            int bsbaCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string course = sh.Range[i, 9].Value.ToString();
                if (course == "BSIT") bsitCourseCount++;
                if (course == "BSED") bsedCourseCount++;
                if (course == "BSBA") bsbaCourseCount++;
            }
            lblBSIT.Text = bsitCourseCount.ToString();
            lblBSED.Text = bsedCourseCount.ToString();
            lblBSBA.Text = bsbaCourseCount.ToString();

           
        }


    
    public void loadform(object Form)
        {
            // create instance
            Form form = Form as Form;
            //configure the form to be loaded
            form.TopLevel = false;
            form.Dock = DockStyle.Fill;

            // Clear existing controls
            this.mainpanel.Controls.Clear();

            //add the form to the panel and display it
            this.mainpanel.Controls.Add(form);
            this.mainpanel.Tag = form;
            form.Show();
        }
        private void btnLogout_Click(object sender, EventArgs e)
        { 

            DialogResult result = MessageBox.Show( "Are you sure you want to log out?", "Confirm Logout",MessageBoxButtons.YesNo, MessageBoxIcon.Question );

            if (result == DialogResult.Yes)
            {
                MyLogs logs = new MyLogs();
                logs.insertLogs(currentUserName, "Successfully logged out!");

                this.Close();
                Login login = new Login();
                login.ShowDialog();
               
                if (result == DialogResult.No)
                {
                    return;
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            loadform(new home());

            // Load the Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            // Count active and inactive students
            int activeStudentCount = 0;
            int inactiveStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string status = sh.Range[i, 13].Value?.ToString().Trim();
                if (status == "1") activeStudentCount++;
                else if (status == "0") inactiveStudentCount++;
            }
            lblActive.Text = activeStudentCount.ToString();
            lblInactive.Text = inactiveStudentCount.ToString();

            // Count the male students
            int maleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }
            }

            // Count the female students
            int femaleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }
            }

            // Count hobbies
            int dancingHobbiesCount = 0;
            int singingHobbiesCount = 0;
            int readingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string hobby = sh.Range[i, 3].Value.ToString();
                if (hobby == "Dancing") dancingHobbiesCount++;
                if (hobby == "Singing") singingHobbiesCount++;
                if (hobby == "Reading") readingHobbiesCount++;
            }
            lblDancing.Text = dancingHobbiesCount.ToString();
            lblSinging.Text = singingHobbiesCount.ToString();
            lblReading.Text = readingHobbiesCount.ToString();

            // Count favorite colors
            int pinkColorCount = 0;
            int blackColorCount = 0;
            int whiteColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string color = sh.Range[i, 5].Value.ToString();
                if (color == "Pink") pinkColorCount++;
                if (color == "Black") blackColorCount++;
                if (color == "White") whiteColorCount++;
            }
            lblPink.Text = pinkColorCount.ToString();
            lblBlack.Text = blackColorCount.ToString();
            lblWhite.Text = whiteColorCount.ToString();

            // Count courses
            int bsitCourseCount = 0;
            int bsedCourseCount = 0;
            int bsbaCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                string course = sh.Range[i, 9].Value.ToString();
                if (course == "BSIT") bsitCourseCount++;
                if (course == "BSED") bsedCourseCount++;
                if (course == "BSBA") bsbaCourseCount++;
            }
            lblBSIT.Text = bsitCourseCount.ToString();
            lblBSED.Text = bsedCourseCount.ToString();
            lblBSBA.Text = bsbaCourseCount.ToString();
           
        }
       
        private void btnActive_Click(object sender, EventArgs e)
        {

            // Prepare the form instance
            Active active = new Active(currentUserName);

            // Load Excel data
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            // Export and filter data
            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                if (dr[12].ToString().Trim() == "1")
                {
                    filtered.ImportRow(dr);
                }
            }
            active.dataGridView1.DataSource = filtered;

            loadform(active);
        }

        private void btnInactive_Click(object sender, EventArgs e)
        {
            // Prepare the form instance
            Inactive inactive = new Inactive(currentUserName);

            // Load Excel data
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            // Export and filter data
            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                if (dr[12].ToString().Trim() == "0")
                {
                    filtered.ImportRow(dr);
                }
            }
            inactive.dataGridView1.DataSource = filtered;

            loadform(inactive);
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            MyLogs logs = new MyLogs();

            Logs logsForm = new Logs();

            // Load Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sheet = book.Worksheets[1]; // Sheet2 for logs

            // Export and filter data
            DataTable dt = sheet.ExportDataTable();
            DataTable filtered = dt.Clone();

            foreach (DataRow dr in dt.Rows)
            {
                // No filtering needed here, just copy all rows for logs
                filtered.ImportRow(dr);
            }

            logsForm.dataGridView2.DataSource = filtered;

            loadform(logsForm); 
        }
    }
}
