using System;
using Microsoft.Office.Interop.Excel;

namespace Excelling
{
    public abstract class Excel : IDisposable
    {
        protected Application _app;
        protected Worksheet _sheet;

        protected Excel()
        {
            _app = new Application();
        }

        public void Dispose()
        {
            if (_app != null)
            {
                if (_app.Workbooks.Count > 0)
                {
                    _app.Workbooks.Close();
                }
                _app.Quit();
            }
        }
    }
}
