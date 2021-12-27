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
            Infos = new List<string>();
        }

        public int Order { get; set; }
        public Teilnehmer Hundefuehrer { get; set; }
        public Teilnehmer Mitte { get; set; }
        public List<Teilnehmer> Figuranten { get; set; }
        public List<string> Infos { get; set; }
    }
}
