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
        public ExcelReader(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _filePath = path;
        }

        public TeachPlan GetPlan()
        {
            using var package = new ExcelPackage(new FileInfo(_filePath));

            var ws = package.Workbook.Worksheets["Титул"];

            Regex trainDirRegex = new(@"Направление подготовки\s*(.*)");
            string trainDir = trainDirRegex.Match(GetPlainText(ws, 29, 4)).Groups[1].Value.Trim();
            string prof = GetPlainText(ws, 30, 4);

            Regex qualRegex = new(@"Квалификация:\s*(.*)");
            string qual = qualRegex.Match(GetPlainText(ws, 40, 3)).Groups[1].Value.Trim();

            return new(trainDir, prof, qual, 
                       GetDisciplines(package),
                       GetCompDescriptions(package));
        }

        static private HashSet<int> DivideNumbers(string s)
        {
            int CharToDigit(char c) 
            {
                if (c > '9' || c < '0') throw new Exception("Not a number");
                else return c - '0';
            }

            return new(s.Select(c => CharToDigit(c)));
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

        static private string? GetPlainText(ExcelWorksheet ws, int row, int col)
        {
            if (IsCellEmpty(ws, row, col))
                return null;
            else
                return MakeTextPlain(ws.Cells[row, col].GetValue<string>());
        }

        static private double ConvertToDouble(string? s)
        {
            if (s is null)
                return 0;
            else
            {
                IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
                return Double.Parse(s, formatter);
            }
        }

        static private Semester GetSemester(ExcelWorksheet ws, ControlType ct, int semNum, int row)
        {
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
            const int kIndex = 3;
            const int kName = 4;

            const int kExam = 5;
            const int kReport = 6;
            const int kMarkedReport = 7;

            const int kHoursPerZE = 11;

            const int kCompetencies = 84;

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

        static private Dictionary<string, string> GetCompDescriptions(ExcelPackage excelPackage)
        {
            var ws = excelPackage.Workbook.Worksheets["Компетенции"];
            var dict = new Dictionary<string, string>();

            const int firstRow = 3;
            const int numberColumn = 2;
            const int indexColumn = 3;
            const int descriptionColumn = 4;


            for (int row = firstRow; ; row++)
            {
                if (IsCellEmpty(ws, row, indexColumn))
                    break;

                else if (ws.Cells[row, numberColumn].Merge) 
                {
                    dict.Add(GetPlainText(ws, row, numberColumn), GetPlainText(ws, row, descriptionColumn));
                }
            }

            return dict;
        }

        private string _filePath;
    }
}
