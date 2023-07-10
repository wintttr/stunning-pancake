using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    class ExcelReader
    {
        private string _filePath;

        public ExcelReader(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _filePath = path;
        }

        private string GetPlainText(ExcelWorksheet ws, string address)
        {
            StringBuilder result = new(ws.Cells[address].GetValue<string>());
            result.Replace("_x000d_", null);
            return result.ToString().Trim();
        }

        public TeachPlan GetPlan()
        {
            string trainDir, prof, qual;
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var ws = package.Workbook.Worksheets.First(t => t.Name == "Титул");
                trainDir = GetPlainText(ws, "D29");
                prof = GetPlainText(ws, "D30");
                qual = GetPlainText(ws, "C40");
            }

            return new(trainDir, prof, qual); 
        }
    }
}
