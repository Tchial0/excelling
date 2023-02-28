using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Excelling
{
    public class ExcelWriter : Excel
    {
        public ExcelWriter() : base()
        {
            Workbook wb = _app.Workbooks.Add();
            _sheet = wb.Worksheets.Add();
        }

        public string this[int row, int column]
        {
            set
            {
                _sheet.Cells[row, column].Value = value;
            }
        }

        public void WriteColumn(int column, IEnumerable<string> values, int startingRow = 1)
        {
            foreach (var value in values)
            {
                _sheet.Cells[startingRow++, column].Value = value;
            }
        }

        public void WriteRow(int row, IEnumerable<string> values, int startingColumn = 1)
        {

            foreach (var value in values)
            {
                _sheet.Cells[row, startingColumn++].Value = value;
            }
        }


        public void Save(string filename)
        {
            if (_sheet != null)
            {
                _sheet.SaveAs(filename);
            }
        }
    }
}
