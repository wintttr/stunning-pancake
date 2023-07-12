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
        private List<string> _competencies;

        public Discipline(string index, string name, ICollection<Semester> semesters, ICollection<string> competencies)
        {
            Index = index;
            Name = name;
            _semesters = new(semesters);
            _competencies = new(competencies);
        }

        public string Name { get; private init; }
        public string Index { get; private init; }
        public double ZETotal { get; init; }
        public double HoursPerZE { get; init; }
        public ReadOnlyCollection<Semester> Semesters => _semesters.AsReadOnly();
        public ReadOnlyCollection<string> Competencies => _competencies.AsReadOnly();
    }
}
