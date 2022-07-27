using System;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace PentruPomagaDeLaNokia
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        string sFileName;
        int iRow;

        public Form1()
        {
            InitializeComponent();
        }

        private void SelectExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog1.Title = "Excel File to Edit";
            OpenFileDialog1.FileName = "";
            OpenFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = OpenFileDialog1.FileName;

                if (sFileName.Trim() != "")
                {
                    readExcel(sFileName);               // READ EXCEL DATA.
                }
            }
        }

        private void readExcel(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"]; // NAME OF THE SHEET.

            Excel.Range userRange = xlWorkSheet.UsedRange;
            int countRecords = userRange.Rows.Count;

            for (iRow = 3; iRow <= countRecords; iRow++)
            {
                string value = xlWorkSheet.Cells[iRow, ColumnRead.Value].Value;
                xlWorkSheet.Cells[iRow, ColumnWrite.Value] = TransformCell(value);
            }

            xlWorkBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
        }

        private void TransformExcel_Click(object sender, EventArgs e)
        {

        }

        private string TransformCell(string data)
        {
            data.Replace(".", "");

            while (data.Contains("(") == true)
            {
                var startIndex = data.IndexOf("(");
                var endIndex = data.IndexOf(")");

                data = data.Remove(startIndex, endIndex - startIndex + 1);
            }

            data = data.Replace("!=False", "TRUE");
            data = data.Replace("=False", "FALSE");
            data = data.Replace("Subscriber", "Sub");
            data = data.Replace("Consumer", "Con");
            data = data.Replace("Flag", "FLAG");
            data = data.Replace("Business", "Bus");
            data = data.Replace(".", "");
            data = data.Replace("-", "");
            data = data.Replace("/", "");
            data = data.Replace("Not", "");
            data = data.Replace("=", " ");
            data = data.Replace("Mbn", "MBN");
            data = data.Replace("Voicemail", "VOICEMAIL");
            data = data.Replace("true", "TRUE");

            data = Regex.Replace(data, @"\s+", " ");

            List<string> arrayString = data.Split(' ').ToList();

            int count = 0;
            for(var i = 0; i < arrayString.Count; i++)
            {
                if (arrayString[i].EndsWith(" on")
                    || arrayString[i].EndsWith("ing")
                    || arrayString[i].Equals("a")
                    || arrayString[i].Equals("previously")
                    || arrayString[i].Equals("created")
                    || (arrayString[i].StartsWith('(') && arrayString[i].EndsWith(')'))
                    || arrayString[i].Equals("for")
                    || arrayString[i].Equals("to")
                    || arrayString[i].Equals("test")
                    || arrayString[i].Equals("data")
                    || arrayString[i].Equals("with")
                    || arrayString[i].Equals("set")
                    || arrayString[i].Equals("by")
                    || arrayString[i].Equals("addon")
                    || arrayString[i].Equals("med")
                    || arrayString[i].Equals("on")
                    || arrayString[i].Equals("from")
                    || arrayString[i].Equals("and"))
                {
                    arrayString.RemoveAt(i);
                    i--;
                }

                if (arrayString[i].Equals("VOICETWIN")) count++;

                if(count == 2)
                {
                    arrayString.RemoveAt(i);
                    i--;
                    count = 0;
                } 
            }

            data = String.Join(" ", arrayString);

            data = Regex.Replace(data, @"\s+", " ");

            var finalResult = data.Replace(" ", "_");

            if(finalResult.EndsWith("_")) finalResult = finalResult.Substring(0, finalResult.Length - 1);

            return finalResult;
        }

        private void ColumnRead_ValueChanged(object sender, EventArgs e)
        {
            if(ColumnRead.Value < 0) return;
            if(ColumnWrite.Value < 0) return;

            SelectExcel.Enabled = true;
        }

        private void ColumnWrite_ValueChanged(object sender, EventArgs e)
        {
            if (ColumnRead.Value < 0) return;
            if (ColumnWrite.Value < 0) return;

            SelectExcel.Enabled = true;
        }
    }
}