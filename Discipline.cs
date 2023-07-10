using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class Discipline
    {
        public Discipline(string name)
        {
            Name = name;
        }

        public string Name { get; private init; }
        public List<Semester> Semesters { get; } = new();
    }
}
