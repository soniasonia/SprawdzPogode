using Microsoft.Office.Interop.Excel;

namespace SprawdzPogode.Handlers
{
    public class ExcelHandler : IOutputHandler
    {
        public Application Excel { get; set; }
        public string Path { get; set; }

        public ExcelHandler(string path, Application excel)
        {
            Excel = excel;
            Path = path;
        }

        public void Start()
        {
            Excel.Visible = true;
            Excel.ScreenUpdating = false;
            Excel.EnableEvents = false;
            Excel.DisplayAlerts = false;
        }

        public void Handle(string[] values)
        {
            Workbook wb = Excel.Workbooks.Open(Path);
            Worksheet ws = Excel.ActiveSheet;
            Excel.Calculation = XlCalculation.xlCalculationManual;
            int row = ws.UsedRange.Rows.Count + 1;
            int cols = Int32.Parse(values[0]);
            int i = 1;
            while (i < values.Length)
            {
                for (int j = 1; j <= cols; j++)
                {
                    ws.Cells[row, j] = values[i++];
                }
                row++;
            }

            wb.Save();
        }

        public void Finish()
        {
            Excel.ScreenUpdating = true;
            Excel.EnableEvents = true;
            Excel.DisplayAlerts = true;
            Excel.Quit();
        }
    }
}
