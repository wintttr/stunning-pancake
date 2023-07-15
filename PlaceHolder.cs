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
        private string _filePath;
        private Dictionary<string, string> _placeDict;

        public PlaceHolder(TeachPlan teachPlan, string filePath)
        {
            try
            {
                _teachPlan = teachPlan;
                foreach(Discipline discipline in _teachPlan.Disciplines)
                {
                    OneDisipline(discipline);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибочка в конструкторе!");
            }
        }

        //Заполнение шаблона для одной дисциплины
        private void OneDisipline(Discipline discipline)
        {
            //Создание объекта документа
            Document document = new Document();
            //Загрузка шаблона 
            document.LoadFromFile(_filePath);
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
            _placeDict.Add("ClassRoom", Convert.ToString(discipline.TotalClassroom));
            _placeDict.Add("Lectures", Convert.ToString(discipline.TotalLectures));
            _placeDict.Add("Labs", Convert.ToString(discipline.TotalLabs));
            _placeDict.Add("Seminars", Convert.ToString(discipline.TotalSeminars));
            _placeDict.Add("KSR", Convert.ToString(discipline.TotalKSR));
            _placeDict.Add("IKR", Convert.ToString(discipline.TotalIKR));
            _placeDict.Add("SR", Convert.ToString(discipline.TotalSR));
            _placeDict.Add("CoursWorkK", Convert.ToString(discipline.TotalCourseWork));
            _placeDict.Add("MatDev", Convert.ToString(discipline.TotalMatDev));
            _placeDict.Add("IndivTasks", Convert.ToString(discipline.TotalIndivTasks));
            _placeDict.Add("Essay", Convert.ToString(discipline.TotalEssay));
            _placeDict.Add("CurrentControl", Convert.ToString(discipline.TotalCurrentControl));
            _placeDict.Add("ControlType", HowIsControl(discipline));
            _placeDict.Add("ExamPrep", Convert.ToString(discipline.TotalExamPrep));
            _placeDict.Add("GeneralLabor", Convert.ToString(discipline.TotalGeneralLabor));
            _placeDict.Add("Classroom", Convert.ToString(discipline.TotalClassroom));
            _placeDict.Add("ZE", Convert.ToString(discipline.TotalZE));

            //Тут будет вставка компетенций и дублирование той таблицы 
            //Получение таблицы 
            Table compitionTable = (Table)document.Sections[0].Tables[0];
            //Добавить 
            for(int i = 1; i < discipline.Competencies.Count+1; i++)//0вая строка это шапка
            {
                compitionTable.AddRow();
                compitionTable.Rows[i].Cells[0].AddParagraph().AppendText(discipline.Competencies[i]);
                compitionTable.Rows[i].Cells[1].AddParagraph().AppendText(_teachPlan.CompDisc[discipline.Competencies[i]]);
                compitionTable.ApplyHorizontalMerge(i, 0, 1);
            }

            //Замена тегов на значение в шаблоне
            foreach (String key in _placeDict.Keys)
            {
                document.Replace(key, _placeDict[key], false, true);
            }
            document.SaveToFile(_filePath+$"\\{discipline.Name}.docx");
            document.Close();
        }

        private void SemesterFill(int num, Semester semester)
        {
            _placeDict.Add($"Num{num}", Convert.ToString(semester.Num));
            _placeDict.Add($"Lectures{num}", Convert.ToString(semester.Lectures));
            _placeDict.Add($"Labs{num}", Convert.ToString(semester.Labs));
            _placeDict.Add($"Seminars{num}", Convert.ToString(semester.Seminars));
            _placeDict.Add($"KSR{num}", Convert.ToString(semester.KSR));
            _placeDict.Add($"IKR{num}", Convert.ToString(semester.IKR));
            _placeDict.Add($"SR{num}", Convert.ToString(semester.SR));
            _placeDict.Add($"CoursWorkK{num}", Convert.ToString(semester.CourseWork));
            _placeDict.Add($"MatDev{num}", Convert.ToString(semester.MatDev));
            _placeDict.Add($"IndivTasks{num}", Convert.ToString(semester.IndivTasks));
            _placeDict.Add($"Essay{num}", Convert.ToString(semester.Essay));
            _placeDict.Add($"CurrentControl{num}", Convert.ToString(semester.CurrentControl));
            _placeDict.Add($"ControlType{num}", semester.Control == ControlType.Exam ? "экзамен" : "зачет");
            _placeDict.Add($"ExamPrep{num}", Convert.ToString(semester.ExamPrep));
            _placeDict.Add($"GeneralLabor{num}", Convert.ToString(semester.GeneralLabor));
            _placeDict.Add($"Classroom{num}", Convert.ToString(semester.Classroom));
            _placeDict.Add($"ZE{num}", Convert.ToString(semester.ZE));

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
            _placeDict.Add($"CoursWorkK{num}", "-");
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
