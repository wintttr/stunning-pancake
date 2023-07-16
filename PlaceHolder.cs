using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;


namespace WinFormsApp1
{

    public class PlaceHolder
    {
        private TeachPlan _teachPlan;
        private IDictionary<string, string> _compDisc;

        private readonly string _patternPath = "pattern.docx";

        private string _filePath;

        private Dictionary<string, string> _placeDict;

        public PlaceHolder(TeachPlan teachPlan, IDictionary<string, string> compDisc, string filePath)
        {
            _teachPlan = teachPlan;
            _compDisc = compDisc;
            _filePath = filePath;
        }

        //Заполнение шаблона для одной дисциплины
        public void OneDiscipline(Discipline discipline)
        {
            //Создание объекта документа
            Document document = new Document();

            //Загрузка шаблона 
            document.LoadFromFile(_patternPath);

            //Создание словоря для замены 
            _placeDict = new Dictionary<string, string>()
            {
                {"TrainingDirection", _teachPlan.TrainingDirection},
                {"Profile", _teachPlan.Profile},
                {"Qualification", _teachPlan.Qualification},
                {"Name", discipline.Name},
                {"index", Convert.ToString(discipline.Index)},
                {"ZETotal", Convert.ToString(discipline.TotalZE)},
                {"HoursPerZE", Convert.ToString(discipline.HoursPerZE)}
            };

            //Заполнение таблицы
            for(int i=1; i < 5; i++)
            {
                if(i-1 < discipline.Semesters.Count)
                {
                    SemesterFill(i, discipline.Semesters[i - 1]);
                }
                else
                {
                    EmptySemester(i);
                }
            }
            //Заполнение первой колонки (итог по семестрам)
            _placeDict.Add("ClassRoom", FormatDouble(discipline.TotalClassroom));
            _placeDict.Add("Lectures", FormatDouble(discipline.TotalLectures));
            _placeDict.Add("Labs", FormatDouble(discipline.TotalLabs));
            _placeDict.Add("Seminars", FormatDouble(discipline.TotalSeminars));
            _placeDict.Add("KSR", FormatDouble(discipline.TotalKSR));
            _placeDict.Add("IKR", FormatDouble(discipline.TotalIKR));
            _placeDict.Add("SR", FormatDouble(discipline.TotalSR));
            _placeDict.Add("CourseWork", FormatDouble(discipline.TotalCourseWork));
            _placeDict.Add("MatDev", FormatDouble(discipline.TotalMatDev));
            _placeDict.Add("IndivTasks", FormatDouble(discipline.TotalIndivTasks));
            _placeDict.Add("Essay", FormatDouble(discipline.TotalEssay));
            _placeDict.Add("CurrentControl", FormatDouble(discipline.TotalCurrentControl));
            _placeDict.Add("ControlType", HowIsControl(discipline));
            _placeDict.Add("ExamPrep", FormatDouble(discipline.TotalExamPrep));
            _placeDict.Add("GeneralLabor", FormatDouble(discipline.TotalGeneralLabor));
            _placeDict.Add("Classroom", FormatDouble(discipline.TotalClassroom));
            _placeDict.Add("ZE", FormatDouble(discipline.TotalZE));

            //Тут будет вставка компетенций и дублирование той таблицы 
            //Получение таблицы 
            Table compitionTable = (Table)document.Sections[0].Tables[0];

            //Добавить 
            for(int i = 0; i < discipline.Competencies.Count; i++)//0вая строка это шапка
            {
                compitionTable.AddRow();
                compitionTable.Rows[i + 1].Cells[0].AddParagraph().AppendText(discipline.Competencies[i]);
                compitionTable.Rows[i + 1].Cells[1].AddParagraph().AppendText(_compDisc[discipline.Competencies[i]]);
                compitionTable.ApplyHorizontalMerge(i, 0, 1);
            }

            //Замена тегов на значение в шаблоне
            foreach (String key in _placeDict.Keys)
            {
                document.Replace($"#{key}#", _placeDict[key], false, true);
            }

            document.SaveToFile(_filePath+$"\\{discipline.Name}.docx");
            document.Close();
        }

        private void SemesterFill(int num, Semester semester)
        {
            _placeDict.Add($"Num{num}", FormatDouble(semester.Num));
            _placeDict.Add($"Lectures{num}", FormatDouble(semester.Lectures));
            _placeDict.Add($"Labs{num}", FormatDouble(semester.Labs));
            _placeDict.Add($"Seminars{num}", FormatDouble(semester.Seminars));
            _placeDict.Add($"KSR{num}", FormatDouble(semester.KSR));
            _placeDict.Add($"IKR{num}", FormatDouble(semester.IKR));
            _placeDict.Add($"SR{num}", FormatDouble(semester.SR));
            _placeDict.Add($"CourseWork{num}", FormatDouble(semester.CourseWork));
            _placeDict.Add($"MatDev{num}", FormatDouble(semester.MatDev));
            _placeDict.Add($"IndivTasks{num}", FormatDouble(semester.IndivTasks));
            _placeDict.Add($"Essay{num}", FormatDouble(semester.Essay));
            _placeDict.Add($"CurrentControl{num}", FormatDouble(semester.CurrentControl));
            _placeDict.Add($"ControlType{num}", semester.Control == ControlType.Exam ? "экзамен" : "зачет");
            _placeDict.Add($"ExamPrep{num}", FormatDouble(semester.ExamPrep));
            _placeDict.Add($"GeneralLabor{num}", FormatDouble(semester.GeneralLabor));
            _placeDict.Add($"Classroom{num}", FormatDouble(semester.Classroom));
            _placeDict.Add($"ZE{num}", FormatDouble(semester.ZE));
        }
        private void EmptySemester(int num)
        {
            _placeDict.Add($"Num{num}", "-");
            _placeDict.Add($"Lectures{num}", "-");
            _placeDict.Add($"Labs{num}", "-");
            _placeDict.Add($"Seminars{num}", "-");
            _placeDict.Add($"KSR{num}", "-");
            _placeDict.Add($"IKR{num}", "-");
            _placeDict.Add($"SR{num}", "-");
            _placeDict.Add($"CourseWork{num}", "-");
            _placeDict.Add($"MatDev{num}","-");
            _placeDict.Add($"IndivTasks{num}", "-");
            _placeDict.Add($"Essay{num}", "-");
            _placeDict.Add($"CurrentControl{num}","-");
            _placeDict.Add($"ControlType{num}", "-");
            _placeDict.Add($"ExamPrep{num}", "-");
            _placeDict.Add($"GeneralLabor{num}", "-");
            _placeDict.Add($"Classroom{num}", "-");
            _placeDict.Add($"ZE{num}", "-");
        }

        private string FormatDouble(double d)
        {
            return d.ToString("0.##");
        }

        //экзамен или не?
        private string HowIsControl(Discipline discipline)
        {
            foreach(Semester sem in discipline.Semesters)
            {
                if(sem.Control == ControlType.Exam)
                {
                    return "экзамен";
                } 
            }
            return "зачет";
        }
    }
}
