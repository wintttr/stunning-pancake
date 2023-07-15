using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    class ExcelReader
    {
        static readonly int kIndex = 3;
        static readonly int kName = 4;

        static readonly int kExam = 5;
        static readonly int kReport = 6;
        static readonly int kMarkedReport = 7;

        static readonly int kTotalZE = 10;
        static readonly int kHoursPerZE = 11;

        static readonly int kCharCount = 8;
        static readonly int kSemBeg = 18;

        static readonly int kZEOffset = 0;
        static readonly int kLecOffset = 1;
        static readonly int kLabOffset = 2;
        static readonly int kSemOffset = 3;
        static readonly int kKSROffset = 4;
        static readonly int kIKROffset = 5;
        static readonly int kSROffset = 6;
        static readonly int kControlOffset = 7;

        static readonly int kCompetencies = 84;

        public ExcelReader(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _filePath = path;
        }

        public TeachPlan GetPlan()
        {
            using var package = new ExcelPackage(new FileInfo(_filePath));

            string trainDir, prof, qual;

            var ws = package.Workbook.Worksheets["Титул"];

            Regex trainDirRegex = new(@"Направление подготовки\s*(.*)");
            trainDir = trainDirRegex.Match(GetPlainText(ws, 29, 4)).Groups[1].Value.Trim();
            prof = GetPlainText(ws, 30, 4);

            Regex qualRegex = new(@"Квалификация:\s*(.*)");
            qual = qualRegex.Match(GetPlainText(ws, 40, 3)).Groups[1].Value.Trim();

            return new(trainDir, prof, qual, 
                       GetDisciplines(package),
                       GetCompetencies(package));
        }

        static private HashSet<int> DivideNumbers(string s)
        {
            HashSet<int> result = new();

            foreach (char c in s)
            {
                if (c > '9' || c < '0') throw new Exception("Not a number");
                else result.Add(c - '0');
            }

            return result;
        }

        static private bool IsCellEmpty(ExcelWorksheet ws, int row, int col)
        {
            var value = ws.Cells[row, col].Value;
            return value == null || value is string v && v.Trim() == "";
        }

        static private string MakeTextPlain(string s)
        {
            StringBuilder result = new(s);
            result.Replace("_x000d_", null);
            return result.ToString().Trim();
        }

        static private string GetPlainText(ExcelWorksheet ws, int row, int col)
        {
            if (IsCellEmpty(ws, row, col))
                return "";
            else
                return MakeTextPlain(ws.Cells[row, col].GetValue<string>());
        }

        static private double ConvertToDouble(string s)
        {
            if (s == "")
                return 0;
            else
            {
                IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
                return Double.Parse(s, formatter);
            }
        }

        static private Semester GetSemester(ExcelWorksheet ws, ControlType ct, int semNum, int row)
        {
            int GetCharCol(int offset)
            {
                return kSemBeg + kCharCount * (semNum - 1) + offset;
            }

            double GetSemChar(int offset)
            {
                var value = GetPlainText(ws, row, GetCharCol(offset));

                return ConvertToDouble(value);
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

        static private List<Discipline> GetDisciplines(ExcelPackage excelPackage)
        {
            var ws = excelPackage.Workbook.Worksheets["План"];

            List<Discipline> disciplines = new();

            const int firstRow = 5;

            for(int row = firstRow; ; row++)
            {
                if (ws.Cells[row, 3].Merge)
                    continue;

                if (IsCellEmpty(ws, row, 3))
                    break;

                string index = GetPlainText(ws, row, kIndex);
                string name = GetPlainText(ws, row, kName);

                double hoursPerZe = ConvertToDouble(GetPlainText(ws, row, kHoursPerZE));

                var exams = DivideNumbers(GetPlainText(ws, row, kExam));
                var reports = DivideNumbers(GetPlainText(ws, row, kReport));

                reports.UnionWith(DivideNumbers(GetPlainText(ws, row, kMarkedReport)));
                reports.ExceptWith(exams);

                List<Semester> semesters = new();
                foreach (int semNum in exams)
                    semesters.Add(GetSemester(ws, ControlType.Exam, semNum, row));

                foreach (int semNum in reports)
                    semesters.Add(GetSemester(ws, ControlType.Report, semNum, row));

                List<string> competencies = new(GetPlainText(ws, row, kCompetencies).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

                disciplines.Add(new(index, name, semesters, competencies) {
                    HoursPerZE = hoursPerZe
                });
            }

            return disciplines;
        }

        static private Dictionary<string, string> GetCompetencies(ExcelPackage excelPackage)
        {
            var ws = excelPackage.Workbook.Worksheets["Компетенции"];

            var dict = new Dictionary<string, string>();

            for (int row = 3; ; row++)
            {
                if (IsCellEmpty(ws, row, 3))
                    break;

                else if (ws.Cells[row, 2].Merge) 
                {
                    dict.Add(GetPlainText(ws, row, 2), GetPlainText(ws, row, 4));
                }
            }

            return dict;
        }

        private string _filePath;
    }
}
