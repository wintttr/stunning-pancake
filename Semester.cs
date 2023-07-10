using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{

    public enum ControlType { Credit, Exam };
    public class Semester
    {
        public Semester(int semesterNumber) 
        {
            Num = semesterNumber;    
        }

        public int Num { get; init; }

        // Аудиторные занятия
        public int Lectures { get; init; }
        public int Labs { get; init; }
        public int Seminars { get; init; }

        // Иная контактная работа
        public int KSR { get; init; }
        public int IKR { get; init; }

        // Самостоятельная работа
        public int CourseWork { get; init; }
        public int MatDev { get; init; }
        public int IndivTasks { get; init; }
        public int Essay { get; init; }
        public int CurrentControl { get; init; }

        // Контроль
        public ControlType Control { get; init; }
        public int ExamPrep { get; init; }
    }
}
