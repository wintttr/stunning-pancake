using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class Discipline
    {
        public string Name { get; private init; }
        public string Index { get; private init; }

        public double TotalGeneralLabor => Semesters.Sum(sem => sem.GeneralLabor);
        public double TotalClassroom => Semesters.Sum(sem => sem.Classroom);
        public double TotalLectures => Semesters.Sum(sem => sem.Lectures);
        public double TotalLabs => Semesters.Sum(sem => sem.Labs);
        public double TotalSeminars => Semesters.Sum(sem => sem.Seminars);
        public double TotalKSR => Semesters.Sum(sem => sem.KSR);
        public double TotalIKR => Semesters.Sum(sem => sem.IKR);
        public double TotalSR => Semesters.Sum(sem => sem.SR);
        public double TotalCourseWork => Semesters.Sum(sem => sem.CourseWork);
        public double TotalMatDev => Semesters.Sum(sem => sem.MatDev);
        public double TotalIndivTasks => Semesters.Sum(sem => sem.IndivTasks);
        public double TotalEssay => Semesters.Sum(sem => sem.Essay);
        public double TotalCurrentControl => Semesters.Sum(sem => sem.CurrentControl);
        public double TotalExamPrep => Semesters.Sum(sem => sem.ExamPrep);
        public double TotalZE => Semesters.Sum(sem => sem.ZE);
        public double HoursPerZE { get; init; }

        public Discipline(string index, string name, ICollection<Semester> semesters, ICollection<string> competencies)
        {
            Index = index;
            Name = name;
            _semesters = new(semesters);
            _competencies = new(competencies);

            _semesters.Sort((s1, s2) => s1.Num - s2.Num);
        }

        public override string ToString()
        {
            return $"{Index} \"{Name}\"";
        }
        public ReadOnlyCollection<Semester> Semesters => _semesters.AsReadOnly();
        public ReadOnlyCollection<string> Competencies => _competencies.AsReadOnly();

        private List<Semester> _semesters;
        private List<string> _competencies;
    }
}
