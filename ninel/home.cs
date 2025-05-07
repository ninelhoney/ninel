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
    public partial class home : Form
    {
        public home()
        {
            InitializeComponent();
           
            // Load the Excel file to count active students
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];

            int activeStudentCount = 0;

            // Loop through the rows and check for active status (column 13 holds the active status)
            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                // If the status in column 13 is "1" (active)
                if (sh.Range[i, 13].Value.ToString() == "1")
                {
                    activeStudentCount++;
                    lblActive.Text = activeStudentCount.ToString();
                }


            }
            int inactiveStudentCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                // If the status in column 13 is "0" (inactive)
                if (sh.Range[i, 13].Value.ToString() == "0")
                {
                    inactiveStudentCount++;
                    lblInactive.Text = inactiveStudentCount.ToString();
                }

            }
            int maleGenderCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 2].Value.ToString() == "Male")
                {
                    maleGenderCount++;
                    lblMale.Text = maleGenderCount.ToString();
                }

            }
            int femaleGenderCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 2].Value.ToString() == "Female")
                {
                    femaleGenderCount++;
                    lblFemale.Text = femaleGenderCount.ToString();
                }

            }
            int dancingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Dancing")
                {
                    dancingHobbiesCount++;
                    lblDancing.Text = dancingHobbiesCount.ToString();
                }

            }
            int singingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Singing")
                {
                    singingHobbiesCount++;
                    lblSinging.Text = singingHobbiesCount.ToString();
                }

            }
            int readingHobbiesCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 3].Value.ToString() == "Reading")
                {
                    readingHobbiesCount++;
                    lblReading.Text = readingHobbiesCount.ToString();
                }

            }
            int pinkColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Pink")
                {
                    pinkColorCount++;
                    lblPink.Text = pinkColorCount.ToString();
                }

            }
            int blackColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "Black")
                {
                    blackColorCount++;
                    lblBlack.Text = blackColorCount.ToString();
                }

            }
            int whiteColorCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 5].Value.ToString() == "White")
                {
                    whiteColorCount++;
                    lblWhite.Text = whiteColorCount.ToString();
                }

            }
            int bsitCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSIT")
                {
                    bsitCourseCount++;
                    lblBSIT.Text = bsitCourseCount.ToString();
                }

            }
            int bsedCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSED")
                {
                    bsedCourseCount++;
                    lblBSED.Text = bsedCourseCount.ToString();
                }

            }
            int bsbaCourseCount = 0;

            for (int i = 2; i <= sh.LastRow; i++) // Start from row 2 to skip header
            {
                if (sh.Range[i, 9].Value.ToString() == "BSBA")
                {
                    bsbaCourseCount++;
                    lblBSBA.Text = bsbaCourseCount.ToString();
                }
            }
        }

        private void home_Load(object sender, EventArgs e)
        {

        }
    }
}
