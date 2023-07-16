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
        public TeachPlan(string trainDir, string prof, string qual)
        {
            TrainingDirection = trainDir;
            Profile = prof;
            Qualification = qual;
        }
        public string TrainingDirection { get; private init; }
        public string Profile { get; private init; }
        public string Qualification { get; private init; }
    }
}
