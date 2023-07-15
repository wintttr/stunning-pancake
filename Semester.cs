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
        public double Classroom => Lectures + Labs + Seminars + KSR + IKR;
        public double Lectures { get; init; } = 0;
        public double Labs { get; init; } = 0;
        public double Seminars { get; init; } = 0;

        // Иная контактная работа
        public double KSR { get; init; } = 0;
        public double IKR { get; init; } = 0;

        // Самостоятельная работа
        public double SR { get; init; } = 0; // НЕ Временно
        public double CourseWork = 0;
        public double MatDev => (SR - CurrentControl) / 2;
        public double IndivTasks => (SR - CurrentControl) / 2;
        public double Essay = 0;
        public double CurrentControl => 2 + (SR - Math.Floor(SR));

        // ЗЕ, зачетные единицы
        public double ZE { get; init; } = 0;

        // Контроль
        public ControlType Control { get; init; }
        public double ExamPrep { get; init; } = 0;

        // Общая трудоёмкость
        public double GeneralLabor => Classroom + SR;
    }
}
