using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    class ExcelReader
    {
        const int kIndex = 3;
        const int kName = 4;

        const int kExam = 5;
        const int kReport = 6;
        const int kMarkedReport = 7;

        const int kCharCount = 8;
        const int kSemBeg = 18;

        const int kZEOffset = 0;
        const int kLecOffset = 1;
        const int kLabOffset = 2;
        const int kSemOffset = 3;
        const int kKSROffset = 4;
        const int kIKROffset = 5;
        const int kSROffset = 6;
        const int kControlOffset = 7;

        private string _filePath;

        public ExcelReader(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _filePath = path;
        }
    
        private string MakeTextPlain(string s)
        {
            StringBuilder result = new(s);
            result.Replace("_x000d_", null);
            return result.ToString().Trim();
        }

        private string GetPlainText(ExcelWorksheet ws, int row, int col)
        {
            return MakeTextPlain(ws.Cells[row, col].GetValue<string>());
        }

        private string GetPlainText(ExcelWorksheet ws, string address)
        {
            return MakeTextPlain(ws.Cells[address].GetValue<string>());
        }

        public TeachPlan GetPlan()
        {
            string trainDir, prof, qual;
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var ws = package.Workbook.Worksheets["Титул"];
                trainDir = GetPlainText(ws, "D29");
                prof = GetPlainText(ws, "D30");
                qual = GetPlainText(ws, "C40");
            }

            return new(trainDir, prof, qual); 
        }

        static private ICollection<int> DivideNumbers(string s)
        {
            HashSet<int> result = new();

            foreach(char c in s)
            {
                if (c > '9' || c < '0') throw new Exception("Not a number");
                else result.Add(c - '0');
            }

            return result;
        }

        public Semester GetSemester(ExcelWorksheet ws, int semNum, int row)
        {
            ControlType ct;

            if (DivideNumbers(GetPlainText(ws, row, kExam)).Contains(semNum))
            {
                ct = ControlType.Exam;
            }
            else if (DivideNumbers(GetPlainText(ws, row, kReport)).Contains(semNum))
            {
                ct = ControlType.Report;
            }
            else if (DivideNumbers(GetPlainText(ws, row, kMarkedReport)).Contains(semNum))
            {
                ct = ControlType.Report;
            }
            else
                throw new Exception("Semester number doesn't exists");

            return new(semNum);
        }

        public ICollection<Discipline> GetDisciplines()
        {
            List<Discipline> disciplines = new();

            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var ws = package.Workbook.Worksheets["План"];

                const int firstRow = 5;

                for(int row = firstRow; ; row++)
                {
                    if (ws.Cells[row, 3].Merge)
                        continue;

                    if (ws.Cells[row, 3].GetValue<string>() == null)
                        break;

                    disciplines.Add(new(GetPlainText(ws, row, kName)));
                }
            }

            return disciplines;
        }
    }
}
