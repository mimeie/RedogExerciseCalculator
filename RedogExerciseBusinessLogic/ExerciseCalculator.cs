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


        }

        public void addTeilnehmer(string name, bool isMitHund)
        {
            Teilnehmer tn = new Teilnehmer();
            tn.Name = name;
            tn.IsMitHund = isMitHund;
            teilnehmerList.Add(tn);
        }

        public void Execute()
        {
            //Teilnehmer Liste durchschütteln          
            teilnehmerList = teilnehmerList.OrderBy(x => Guid.NewGuid()).ToList();

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

            //figurantenplätze besetzen
            int figurantPlatz = 1;

            foreach (Teilnehmer figurant in teilnehmerList)
            {
                if (figurantPlatz  > calcSetting.AnzahlFiguranten)
                {
                    figurantPlatz = 1;
                }
                figurant.FigurantenPlatz = figurantPlatz;

                figurantPlatz++;
            }

            //mit der geforderten Anzahl Figuranten besetzen
            for (figurantPlatz = 1; figurantPlatz <= calcSetting.AnzahlFiguranten; figurantPlatz++)
            {
                
                List<Teilnehmer> figuranten = new List<Teilnehmer>();                
                foreach (Uebungsrunde runde in uebungsplan.ToList().OrderBy(x => x.Order))
                {
                    //alle möglichen Figuranten laden 
                    figuranten.AddRange(teilnehmerList.Where(x=> x.FigurantenPlatz == figurantPlatz).ToList());

                    //Dann rausfiltern filtern (nicht der nächste HF, nicht der jetzige Hf)
                    figuranten.RemoveAll(x => x == runde.Hundefuehrer);                    
                    if (uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault() != null)
                    { 
                        figuranten.RemoveAll(x => x == uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault().Hundefuehrer);
                    }

                    //letzten Figurant laden für jeweilige Position
                    runde.Figuranten.Add(figuranten.FirstOrDefault());

                }
            }




            


        }
    }
}
