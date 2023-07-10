using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class TeachPlan
    {
        public TeachPlan(string name, string trainDir, string prof, string qual)
        {
            Name = name;
            TrainingDirection = trainDir;
            Profile = prof;
            Qualification = qual;
        }

        public string Name { get; private init; }
        public string TrainingDirection { get; private init; }
        public string Profile { get; private init; }
        public string Qualification { get; private init; }

        public List<Discipline> Disciplines { get; } = new();
    }
}
