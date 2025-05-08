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

namespace ninel
{
    public partial class Logs : Form
    {
        public Logs()
        {
            InitializeComponent();
   
        }
        
       
        private void btnSearch_Click(object sender, EventArgs e)
        {
            //i removed the break 
            dataGridView2.ClearSelection();
            bool itemFound = false;

            try
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
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
                            dataGridView2.FirstDisplayedScrollingRowIndex = row.Index;
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
    }
}
