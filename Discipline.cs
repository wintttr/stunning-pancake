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
        private List<Semester> _semesters;

        public Discipline(string index, string name, ICollection<Semester>? semesters = null)
        {
            Index = index;
            Name = name;

            if (semesters == null)
                _semesters = new();
            else
                _semesters = new(semesters);
        }

        public string Name { get; private init; }
        public string Index { get; private init; }
        public ReadOnlyCollection<Semester> Semesters => _semesters.AsReadOnly();

        public void Add(Semester s)
        {
            _semesters.Add(s);
        }
    }
}
