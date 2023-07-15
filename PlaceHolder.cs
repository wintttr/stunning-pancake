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
                {"ZETotal", Convert.ToString(discipline.ZETotal)},
                {"HoursPerZE", Convert.ToString(discipline.HoursPerZE)}
            };
            //Заполнение таблицы (надо доделать общую трудоемкость)
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
            //_placeDict.Add($"Num", Convert.ToString());
            //_placeDict.Add($"Lectures", Convert.ToString());
            //_placeDict.Add($"Labs", Convert.ToString());
            //_placeDict.Add($"Seminars", Convert.ToString());
            //_placeDict.Add($"KSR", Convert.ToString());
            //_placeDict.Add($"IKR", Convert.ToString());
            //_placeDict.Add($"SR", Convert.ToString());
            //_placeDict.Add($"CoursWorkK", Convert.ToString());
            //_placeDict.Add($"MatDev", Convert.ToString());
            //_placeDict.Add($"IndivTasks", Convert.ToString());
            //_placeDict.Add($"Essay", Convert.ToString());
            //_placeDict.Add($"CurrentControl", Convert.ToString());
            //_placeDict.Add($"ControlType{num}", semester.Control);
            //_placeDict.Add($"ExamPrep", Convert.ToString());

            //Тут будет вставка компетенций и дублирование той таблицы 

            //Замена тегов на значение в шаблоне
            foreach(String key in _placeDict.Keys)
            {
                document.Replace(key, _placeDict[key], false, true);
            }
            document.SaveToFile(_filePath);
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
            //_placeDict.Add($"ControlType{num}", semester.Control);
            _placeDict.Add($"ExamPrep{num}", Convert.ToString(semester.ExamPrep));
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
        }
    }
}
