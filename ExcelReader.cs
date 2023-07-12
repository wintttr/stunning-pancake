using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
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
    
        private bool IsCellEmpty(ExcelWorksheet ws, int row, int col)
        {
            var value = ws.Cells[row, col].Value;
            return value == null || value is string v && v.Trim() == "";
        }

        private string MakeTextPlain(string s)
        {
            StringBuilder result = new(s);
            result.Replace("_x000d_", null);
            return result.ToString().Trim();
        }

        private string GetPlainText(ExcelWorksheet ws, int row, int col)
        {
            if (IsCellEmpty(ws, row, col))
                return "";
            else
                return MakeTextPlain(ws.Cells[row, col].GetValue<string>());
        }

        public TeachPlan GetPlan()
        {
            string trainDir, prof, qual;
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var ws = package.Workbook.Worksheets["Титул"];
                trainDir = GetPlainText(ws, 29, 4);
                prof = GetPlainText(ws, 30, 4);
                qual = GetPlainText(ws, 40, 3);
            }

            return new(trainDir, prof, qual); 
        }

        static private HashSet<int> DivideNumbers(string s)
        {
            HashSet<int> result = new();

            foreach(char c in s)
            {
                if (c > '9' || c < '0') throw new Exception("Not a number");
                else result.Add(c - '0');
            }

            return result;
        }

        private Semester GetSemester(ExcelWorksheet ws, ControlType ct, int semNum, int row)
        {
            int GetCharCol(int offset)
            {
                return kSemBeg + kCharCount * (semNum - 1) + offset;
            }

            double GetSemChar(int offset)
            {
                var value = GetPlainText(ws, row, GetCharCol(offset));

                if (value == "")
                    return 0;
                else
                {
                    IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
                    return Double.Parse(value, formatter);
                }
            }

            double ZE = GetSemChar(kZEOffset);
            double Lectures = GetSemChar(kLecOffset);
            double Labs = GetSemChar(kLabOffset);
            double Seminars = GetSemChar(kSemOffset);
            double KSR = GetSemChar(kKSROffset);
            double IKR = GetSemChar(kIKROffset);
            double SR = GetSemChar(kSROffset);

            double Control = 0;
            if(ct == ControlType.Exam)
                Control = GetSemChar(kControlOffset);

            return new(semNum)
            {
                ZE = ZE,
                Lectures = Lectures,
                Labs = Labs,
                Seminars = Seminars,
                KSR = KSR,
                IKR = IKR,
                SR = SR,
                Control = ct,
                ExamPrep = Control
            };
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

                    string index = GetPlainText(ws, row, kIndex);
                    string name = GetPlainText(ws, row, kName);

                    List<Semester> semesters = new();

                    var exams = DivideNumbers(GetPlainText(ws, row, kExam));
                    var reports = DivideNumbers(GetPlainText(ws, row, kReport));

                    reports.UnionWith(DivideNumbers(GetPlainText(ws, row, kMarkedReport)));
                    reports.ExceptWith(exams);


                    foreach (int semNum in exams)
                        semesters.Add(GetSemester(ws, ControlType.Exam, semNum, row));

                    foreach (int semNum in reports)
                        semesters.Add(GetSemester(ws, ControlType.Report, semNum, row));

                    disciplines.Add(new(index, name, semesters));
                }
            }

            return disciplines;
        }
    }
}
