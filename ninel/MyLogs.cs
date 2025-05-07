using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ninel
{
    internal class MyLogs
    {
        Workbook book = new Workbook();

        public void insertLogs(string user, string message)
        {
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
            Worksheet sh = book.Worksheets[0];
            int r = sh.Rows.Length + 1;
            sh.Range[r, 1].Value = user;
            sh.Range[r, 2].Value = message;
            sh.Range[r, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[r, 4].Value = DateTime.Now.ToString("hh:mm:ss tt");
            book.LoadFromFile("C:\\Users\\ACT-STUDENT\\Downloads\\newwwww\\ninel(V2)\\Book1.xlsx");
        }
    }
}
