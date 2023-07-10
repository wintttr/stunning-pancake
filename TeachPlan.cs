using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class TeachPlan
    {
        private List<Discipline> _disciplines;

        public TeachPlan(string trainDir, string prof, string qual)
        {
            TrainingDirection = trainDir;
            Profile = prof;
            Qualification = qual;

            _disciplines = new();
        }

        public string TrainingDirection { get; private init; }
        public string Profile { get; private init; }
        public string Qualification { get; private init; }
        public ReadOnlyCollection<Discipline> Disciplines => _disciplines.AsReadOnly();

        public void Add(Discipline d)
        {
            _disciplines.Add(d);
        }
    }
}
