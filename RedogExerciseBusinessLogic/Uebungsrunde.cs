using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RedogExerciseBusinessLogic
{
    public class Uebungsrunde
    {
        public Uebungsrunde()
         {
            Figuranten = new List<Teilnehmer>();
        }

        public int Order { get; set; }
        public Teilnehmer Hundefuehrer { get; set; }
        public Teilnehmer Mitte { get; set; }
        public List<Teilnehmer> Figuranten { get; set; }
        public string Info { get; set; }
    }
}
