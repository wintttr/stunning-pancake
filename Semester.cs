using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public enum ControlType { Report, Exam };
    public class Semester
    {
        public Semester(int semesterNumber) 
        {
            Num = semesterNumber;    
        }

        public int Num { get; private init; }

        // Аудиторные занятия
        public int Lectures { get; init; } = 0;
        public int Labs { get; init; } = 0;
        public int Seminars { get; init; } = 0;

        // Иная контактная работа
        public int KSR { get; init; } = 0;
        public int IKR { get; init; } = 0;

        // Самостоятельная работа

        public int SR { get; init; } = 0; // Временно
        public int CourseWork { get; init; } = 0;
        public int MatDev { get; init; } = 0;
        public int IndivTasks { get; init; } = 0;
        public int Essay { get; init; } = 0;
        public int CurrentControl { get; init; } = 0;

        // ЗЕ, зачетные единицы
        public int ZE { get; init; } = 0;

        // Контроль
        public ControlType Control { get; init; }
        public int ExamPrep { get; init; } = 0;
    }
}
