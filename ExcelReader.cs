using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.CodeDom;
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
    class ParseErrorException : ApplicationException
    {
        public ParseErrorException(string? s) : base(s) { }
    }

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
            Regex qualRegex = new(@"Квалификация:\s*(.*)");

            string rawTrainDir = GetStringFromCell(ws.Cells[29, 4]) ?? 
                                            throw new ParseErrorException("Train direction not found");

            string profile = GetStringFromCell(ws.Cells[30, 4]) ?? 
                                            throw new ParseErrorException("Profile not found");

            string rawQualification = GetStringFromCell(ws.Cells[40, 3]) ?? 
                                            throw new ParseErrorException("Qualification not found");

            string trainDir = MatchToRegex(trainDirRegex, rawTrainDir);
            string qualification = MatchToRegex(qualRegex, rawQualification);

            return new(trainDir, profile, qualification,
                       GetDisciplines(package),
                       GetCompDescriptions(package));
        }

        static private string MatchToRegex(Regex regex, string pattern)
        {
            return regex.Match(pattern).Groups[1].Value.Trim();
        }

        static private HashSet<int> DivideNumbers(string? s)
        {
            static int CharToDigit(char c) 
            {
                if (c > '9' || c < '0') throw new Exception("Not a number");
                else return c - '0';
            }

            if (s is null)
                return new();
            else
                return new(s.Select(c => CharToDigit(c)));
        }

        static private bool IsCellEmpty(ExcelRange cellRange)
        {
            var value = cellRange.Value;
            return value == null || value is string v && v.Trim() == "";
        }

        static private string MakeTextPlain(string s)
        {
            StringBuilder result = new(s);
            result.Replace("_x000d_", null);
            return result.ToString().Trim();
        }

        static private string? GetStringFromCell(ExcelRange cellRange)
        {
            if (IsCellEmpty(cellRange))
                return null;
            else
                return MakeTextPlain(cellRange.GetValue<string>());
        }

        static private double ConvertToDouble(string s)
        {
            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
            return Double.Parse(s, formatter);
        }

        static private double? GetDoubleFromCell(ExcelRange cellRange)
        {
            if (IsCellEmpty(cellRange))
                return null;
            else
                return ConvertToDouble(cellRange.GetValue<string>());
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

            double GetSemChar(int offset) => GetDoubleFromCell(ws.Cells[row, GetCharCol(offset)]) ?? 0;

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

                if (IsCellEmpty(ws.Cells[row, 3]))
                    break;

                string index = GetStringFromCell(ws.Cells[row, kIndex]) ?? 
                                            throw new ParseErrorException("Discipline index not found");

                string name = GetStringFromCell(ws.Cells[row, kName]) ?? 
                                            throw new ParseErrorException("Discipline name not found");

                double hoursPerZe = GetDoubleFromCell(ws.Cells[row, kHoursPerZE]) ?? 
                                            throw new ParseErrorException("Hours per ZE not found");

                var exams = DivideNumbers(GetStringFromCell(ws.Cells[row, kExam]));
                var reports = DivideNumbers(GetStringFromCell(ws.Cells[row, kReport]));

                reports.UnionWith(DivideNumbers(GetStringFromCell(ws.Cells[row, kMarkedReport])));
                reports.ExceptWith(exams);

                List<Semester> semesters = new();
                foreach (int semNum in exams)
                    semesters.Add(GetSemester(ws, ControlType.Exam, semNum, row));

                foreach (int semNum in reports)
                    semesters.Add(GetSemester(ws, ControlType.Report, semNum, row));

                string rawCompetenciesString = GetStringFromCell(ws.Cells[row, kCompetencies]) ?? 
                                            throw new ParseErrorException("Competencies not found");

                List<string> competencies = new(rawCompetenciesString.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

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
                if (IsCellEmpty(ws.Cells[row, indexColumn]))
                    break;

                else if (ws.Cells[row, numberColumn].Merge) 
                {
                    string compNumber = GetStringFromCell(ws.Cells[row, numberColumn]) ?? 
                                                throw new ParseErrorException("Competency number not found");
                    string compDescription = GetStringFromCell(ws.Cells[row, descriptionColumn]) ?? 
                                                throw new ParseErrorException("Competency description not found");

                    dict.Add(compNumber, compDescription);
                }
            }

            return dict;
        }

        private string _filePath;
    }
}
