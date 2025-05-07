using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ninel
{
    class MyLogs
    {
        Workbook book = new Workbook();

        public void insertLogs(string user, string message)
        {
            //logs
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");
            Worksheet sh = book.Worksheets[1];
            int r = sh.LastRow + 1;

            sh.Range[r, 1].Value = user;
            sh.Range[r, 2].Value = message;
            sh.Range[r, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[r, 4].Value = DateTime.Now.ToString("HH:mm:ss");

            book.SaveToFile("C:\\Users\\ninel\\source\\repos\\ninel\\Book1.xlsx");

        }
    }
}
