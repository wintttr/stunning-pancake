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
        public TeachPlan(string trainDir, string prof, string qual, ICollection<Discipline> disciplines, IDictionary<string, string> compDisc)
        {
            TrainingDirection = trainDir;
            Profile = prof;
            Qualification = qual;

            _disciplines = new(disciplines);
            _compDisc = new(compDisc);
        }
        public string TrainingDirection { get; private init; }
        public string Profile { get; private init; }
        public string Qualification { get; private init; }

        public ReadOnlyCollection<Discipline> Disciplines => _disciplines.AsReadOnly();
        public ReadOnlyDictionary<string, string> CompDisc => new(_compDisc);

        private List<Discipline> _disciplines;
        private Dictionary<string, string> _compDisc; 
    }
}
