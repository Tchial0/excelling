using System;
using System.Collections.Generic;

namespace Excelling
{
    public class ExcelReader : Excel
    {
        public ExcelReader(string fileName) : base()
        {
            var wb = _app.Workbooks.Open(fileName);
            _sheet = wb.Worksheets[1];
        }

        public string this[int row, int column]
        {
            get
            {
                string value;
                try
                {
                    value = _sheet.Cells[row, column].Value.ToString();
                }
                catch (Exception)
                {
                    value = string.Empty;
                }
                return value;
            }
        }

        public List<string> GetColumn(int column, int startingRow = 1)
        {
            List<string> values = new List<string>();
            for (int row = startingRow; this[row, column] != string.Empty; row++)
            {
                values.Add(_sheet.Cells[row, column].Value.ToString());
            }
            return values;
        }

        public List<string> GetRow(int row, int startingColumn = 1)
        {
            List<string> values = new List<string>();
            for (int column = startingColumn; this[row, column] != string.Empty; column++)
            {
                values.Add(_sheet.Cells[row, column].Value.ToString());
            }
            return values;
        }

    }
}
