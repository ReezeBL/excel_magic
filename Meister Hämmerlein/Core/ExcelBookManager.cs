using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Meister_Hämmerlein.Core
{
    public class ExcelBookManager : IDisposable
    {
        private readonly Application application = new Application {DisplayAlerts = false};

        public Workbook OpenWorkbook(string path)
        {
            return application.Workbooks.Open(path);
        }

        public Workbook CreateWorkbook()
        {
            return application.Workbooks.Add(System.Reflection.Missing.Value);
        }

        private void ReleaseUnmanagedResources()
        {
            foreach (Workbook workbook in application.Workbooks)
            {
                workbook.Close(true);
                Marshal.ReleaseComObject(workbook);
            }
            application.Quit();
            Marshal.ReleaseComObject(application);
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~ExcelBookManager()
        {
            ReleaseUnmanagedResources();
        }
    }
}