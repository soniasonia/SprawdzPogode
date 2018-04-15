using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using System.Drawing;

namespace SprawdzPogode
{
    class Extractor
    {
        public Extractor(string[] lines)
        {
            this.lines = lines;
        }
        private string[] lines;

        public bool CheckForError(string s1, string s2, string s3)
        {
            if (s1.Equals("[not found]"))
            {
                return true;
            }
            if (s2.Equals("[not found]"))
            {
                return true;
            }
            if (s3.Equals("[not found]"))
            {
                return true;
            }
            return false;
        }

        public void ExtractData(ExcelHandler ex, ChromeHandler ch)
        {
            Worksheet ws = ex.Ws;
            int row = ws.UsedRange.Rows.Count + 1;
            int counter = 0;

            foreach (string line in lines)
            {
                Console.WriteLine("City: " + line);

                ch.Search(line);
                ws.Cells[row, 1].Value = String.Format("{0:yyyy/MM/dd HH:mm:ss}", DateTime.Now);
                ws.Cells[row, 2].Value = line;
                string temp = ch.GetData(By.Id("wob_tm"));
                string rain = ch.GetData(By.Id("wob_pp"));
                string wind = ch.GetData(By.Id("wob_ws"));
                ws.Cells[row, 3].Value = temp;
                ws.Cells[row, 4].Value = rain;
                ws.Cells[row, 5].Value = wind;

                if (CheckForError(temp, rain, wind))
                {
                    ws.Cells[row, 6].Value = "Fail";
                    ws.Range[ws.Cells[row, 1], ws.Cells[row, 5]].Interior.Color = Color.Red;
                }
                else
                {
                    ws.Cells[row, 6].Value = "Success";
                }
                row++;
                counter++;
            }

        }


    }
}
