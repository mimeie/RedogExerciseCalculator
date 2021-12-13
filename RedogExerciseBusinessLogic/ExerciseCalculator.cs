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
        public List<Uebungsrunde> uebungsplan;

        public ExerciseCalculator()
        {
            calcSetting = new CalculatorSettings();
            teilnehmerList = new List<Teilnehmer>();
            uebungsplan = new List<Uebungsrunde>();


            //Daten im Moment hier manuell abfüllen
            calcSetting.AnzahlFiguranten = 2;
            calcSetting.AnzahlRundenDraussen = 2;


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
            //alle HF einplanen
            int i = 0;
            foreach (Teilnehmer tn in teilnehmerList.Where(x => x.IsMitHund == true).ToList())
            {
                i++;
                Uebungsrunde runde = new Uebungsrunde();
                runde.Order = i;
                runde.Hundefuehrer = tn;                

                uebungsplan.Add(runde);
            }

            //mit Figuranten besetzen
            for (int j = 1; j <= calcSetting.AnzahlFiguranten; j++)
            {
                //letzten Figurant laden für jeweilige Position
                List<Teilnehmer> figuranten = new List<Teilnehmer>();
                
                foreach (Uebungsrunde runde in uebungsplan.ToList().OrderBy(x => x.Order))
                {
                    
                }
            }




            //runde.Figuranten.Add(tn);


        }
    }
}
