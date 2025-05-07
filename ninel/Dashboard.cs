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
        public Dashboard(string name, string path)
        {
            InitializeComponent();

            //retrieve the username 
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

            // Load the Excel file
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            //count the active students
            int activeStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++) 
            {
                if (sh.Range[i, 13].Value.ToString() == "1")
                {
                    activeStudentCount++;
                    lblActive.Text = activeStudentCount.ToString();
                }
            }

            //count the inactive students
            int inactiveStudentCount = 0;
            for (int i = 2; i <= sh.LastRow; i++) 
            {
                if (sh.Range[i, 13].Value.ToString() == "0")
                {
                    inactiveStudentCount++;
                    lblInactive.Text = inactiveStudentCount.ToString();
                }
            }

            //count the male students
            int maleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }
            }
            //count the female students
            int femaleGenderCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }
            }

            //count the students who's hobbies is dancing
            int dancingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 3].Value.ToString() == "Dancing")
                {
                    dancingHobbiesCount++;
                    lblDancing.Text = dancingHobbiesCount.ToString();
                }
            }

            //count the students who's hobbies is singing
            int singingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 3].Value.ToString() == "Singing")
                {
                    singingHobbiesCount++;
                    lblSinging.Text = singingHobbiesCount.ToString();
                }
            }

            //count the students who's hobbies is reading
            int readingHobbiesCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 3].Value.ToString() == "Reading")
                {
                    readingHobbiesCount++;
                    lblReading.Text = readingHobbiesCount.ToString();
                }
            }

            //count the students who's favcolor is pink
            int pinkColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 5].Value.ToString() == "Pink")
                {
                    pinkColorCount++;
                    lblPink.Text = pinkColorCount.ToString();
                }
            }

            //count the students who's favcolor is black
            int blackColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 5].Value.ToString() == "Black")
                {
                    blackColorCount++;
                    lblBlack.Text = blackColorCount.ToString();
                }
            }

            //count the students who's favcolor is white
            int whiteColorCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 5].Value.ToString() == "White")
                {
                    whiteColorCount++;
                    lblWhite.Text = whiteColorCount.ToString();
                }
            }

            //count the students who's course is bsit
            int bsitCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 9].Value.ToString() == "BSIT")
                {
                    bsitCourseCount++;
                    lblBSIT.Text = bsitCourseCount.ToString();
                }
            }

            //count the students who's course is bsed
            int bsedCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 9].Value.ToString() == "BSED")
                {
                    bsedCourseCount++;
                    lblBSED.Text = bsedCourseCount.ToString();
                }
            }

            //count the students who's course is bsba
            int bsbaCourseCount = 0;
            for (int i = 2; i <= sh.LastRow; i++)
            {
                if (sh.Range[i, 9].Value.ToString() == "BSBA")
                {
                    bsbaCourseCount++;
                    lblBSBA.Text = bsbaCourseCount.ToString();
                }
            }
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
                // Close or hide main form
                this.Hide(); // or this.Close();

                if (result != DialogResult.Yes)
                {
                    this.Close();
                }
            }

            Login login = new Login();
            login.ShowDialog();

        }
        private void button1_Click(object sender, EventArgs e)
        {
            loadform(new home());
            //test
        }
        private void btnActive_Click(object sender, EventArgs e)
        {
            // Prepare the form instance
            Active active = new Active();

            // Load Excel data
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
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
            Inactive inactive = new Inactive();

            // Load Excel data
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
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
            loadform(new Logs());
        }

    }
}
