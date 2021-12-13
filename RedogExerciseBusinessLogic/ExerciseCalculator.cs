using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RedogExerciseBusinessLogic
{
    public class ExerciseCalculator
    {
        public List<Teilnehmer> teilnehmerList;
        public CalculatorSettings calcSetting;
        public Uebungsplan uebungsplan;

        public ExerciseCalculator()
        {
            calcSetting = new CalculatorSettings();
            teilnehmerList = new List<Teilnehmer>();
            uebungsplan = new Uebungsplan();


            //Daten im Moment hier manuell abfüllen
            calcSetting.AnzahlFiguranten = 2;


            addTeilnehmer("Michael", true);
            addTeilnehmer("Sarah", true);
            addTeilnehmer("Tom", false);
            addTeilnehmer("Alain", true);
            addTeilnehmer("Clemens", true);
            addTeilnehmer("Joli", true);

        }

        private void addTeilnehmer(string name, bool isMitHund)
        {
            Teilnehmer tn = new Teilnehmer();
            tn.Name = name;
            tn.IsMitHund = isMitHund;
            teilnehmerList.Add(tn);
        }

        public void Execute()
        {
            //alle HF sollen ablaufen
            int i = 0;
            foreach (Teilnehmer tn in teilnehmerList.Where(x => x.IsMitHund == true).ToList())
            {
                i++;
                Uebungsrunde runde = new Uebungsrunde();
                runde.Order = i;
                runde.Hundefuehrer = tn;

                uebungsplan.Uebungsrunde.Add(runde);
            }

                
        }
    }
}
