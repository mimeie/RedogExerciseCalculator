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
            Initialize();
        }

        public void Initialize()
        {

            teilnehmerList = new List<Teilnehmer>();
            uebungsplan = new List<Uebungsrunde>();
        }

        public void addTeilnehmer(string name, bool isMitHund, bool isMitte, bool isSpezialprogrammAnfang)
        {
            Teilnehmer tn = new Teilnehmer();
            tn.Name = name;
            tn.IsMitHund = isMitHund;
            tn.IsMitte = isMitte;
            tn.IsSpezialprogrammAnfang = isSpezialprogrammAnfang;
            teilnehmerList.Add(tn);
        }

        public void addTeilnehmer(string name, string isMitHund, string isMitte, string isSpezialprogrammAnfang)
        {

            bool isMitHundLocal = false;
            bool isMitteLocal = false;
            bool isSpezialprogrammAnfangLocal = false;

            if (isMitHund == "x")
            {
                isMitHundLocal = true;
            }

            if (isMitte == "x")
            {
                isMitteLocal = true;
            }
            if (isSpezialprogrammAnfang == "x")
            {
                isSpezialprogrammAnfangLocal = true;
            }

            addTeilnehmer(name, isMitHundLocal, isMitteLocal, isSpezialprogrammAnfangLocal);

        }

        public void Execute()
        {
            //Teilnehmer Liste durchschütteln          
            teilnehmerList = teilnehmerList.OrderBy(x => Guid.NewGuid()).ToList();

            //alle HF einplanen
            int i = 0;
            foreach (Teilnehmer tn in teilnehmerList.Where(x => x.IsMitHund == true).ToList().OrderByDescending(z => z.IsSpezialprogrammAnfang).OrderBy(y => y.IsMitte))
            {
                i++;
                Uebungsrunde runde = new Uebungsrunde();
                runde.Order = i;
                runde.Hundefuehrer = tn;

                uebungsplan.Add(runde);
            }


            //Mitte einplanen, gleichmässig aufteilen der Reihe nach, je nach Typ
            if (calcSetting.KonfigurationMitte != KonfigurationMitte.KeineMitte)
            {
                List<Teilnehmer> mitteList = new List<Teilnehmer>();
                mitteList = teilnehmerList.Where(y => y.IsMitte == true).ToList();
                
                int mitteTeilnehmer = 0;

                if (calcSetting.KonfigurationMitte == KonfigurationMitte.MitteAmStueck)
                {
                    double wechselRunde = uebungsplan.Count / mitteList.Count;
                    wechselRunde = Math.Round(wechselRunde, 0);
                    int mitteOrder = 1;

                    foreach (Uebungsrunde runde in uebungsplan.ToList().OrderBy(x => x.Order))
                    {

                        //mitte geher abwechseln gleichmässig, wenn HF und Mitte identisch, alternative laden (was noch verfügbar ist
                        if (runde.Hundefuehrer == mitteList.ElementAt(mitteTeilnehmer))
                        {
                            List<Teilnehmer> mitteListAlternative = new List<Teilnehmer>();
                            mitteListAlternative = teilnehmerList.Where(y => y.IsMitte == true).ToList();
                            mitteListAlternative.Remove(mitteList.ElementAt(mitteTeilnehmer));
                            runde.Mitte = mitteListAlternative.FirstOrDefault();
                        }
                        else
                        {
                            runde.Mitte = mitteList.ElementAt(mitteTeilnehmer);
                        }
                        mitteOrder++;

                        if (mitteOrder >= wechselRunde && mitteTeilnehmer + 1 < mitteList.Count)
                        {
                            mitteTeilnehmer++;
                            mitteOrder = 1;
                        }
                    }
                }
                else if (calcSetting.KonfigurationMitte == KonfigurationMitte.MitteAbwechselnd)
                {
                    foreach (Uebungsrunde runde in uebungsplan.ToList().OrderBy(x => x.Order))
                    {
                        if (mitteTeilnehmer > mitteList.Count - 1)
                        {
                            mitteTeilnehmer = 0;
                        }
                        if (runde.Hundefuehrer == mitteList.ElementAt(mitteTeilnehmer))
                        {
                            List<Teilnehmer> mitteListAlternative = new List<Teilnehmer>();
                            mitteListAlternative = teilnehmerList.Where(y => y.IsMitte == true).ToList();
                            mitteListAlternative.Remove(mitteList.ElementAt(mitteTeilnehmer));
                            runde.Mitte = mitteListAlternative.FirstOrDefault();
                        }
                        else
                        {
                            runde.Mitte = mitteList.ElementAt(mitteTeilnehmer);
                        }                        
                        mitteTeilnehmer++;
                    }
                }
            }

            //figurantenplätze besetzen
            int figurantPlatz = 1;
            foreach (Teilnehmer figurant in teilnehmerList.OrderBy(x => x.IsMitHund))
            {
                if (figurantPlatz > calcSetting.AnzahlFiguranten)
                {
                    figurantPlatz = 1;
                }
                figurant.FigurantenPlatz = figurantPlatz;

                figurantPlatz++;
            }

            //mit der geforderten Anzahl Figuranten besetzen
            for (figurantPlatz = 1; figurantPlatz <= calcSetting.AnzahlFiguranten; figurantPlatz++)
            {


                foreach (Uebungsrunde runde in uebungsplan.ToList().OrderBy(x => x.Order))
                {
                    List<Teilnehmer> figuranten = new List<Teilnehmer>();
                    List<Teilnehmer> figurantenGeloeschtWegenMaxAnzahl = new List<Teilnehmer>();

                    //alle möglichen Figuranten laden 
                    figuranten.AddRange(teilnehmerList.Where(x => x.FigurantenPlatz == figurantPlatz).ToList());

                    //Dann rausfiltern filtern (nicht der nächste HF, nicht der jetzige Hf, nicht der vorherige HF)
                    figuranten.RemoveAll(x => x == runde.Hundefuehrer);
                    if (uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault() != null)
                    {
                        figuranten.RemoveAll(x => x == uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault().Hundefuehrer);
                    }
                    if (uebungsplan.Where(y => y.Order == runde.Order - 1).FirstOrDefault() != null)
                    {
                        figuranten.RemoveAll(x => x == uebungsplan.Where(y => y.Order == runde.Order - 1).FirstOrDefault().Hundefuehrer);
                    }

                    //Dann momentan die Mitte komplett als Figuranten rauslöschen                    
                    //foreach (Teilnehmer mitte in mitteList)
                    //{ 
                    //    figuranten.Remove(mitte);
                    //}

                    //den Mitte teilnehmer als Figurant löschen
                    figuranten.Remove(runde.Mitte);

                    //rausfiltern wer schon oft genug war
                    foreach (Teilnehmer gewaehlterFigurant in figuranten.ToList())
                    {
                        int figurantFoundCounter = 0;
                        //alle anderen Runden die schon zugewiesen wurden ohne diese
                        foreach (Uebungsrunde andereRunde in uebungsplan.Where(x => x != runde && x.Figuranten.Count != 0).ToList().OrderBy(x => x.Order))
                        {
                            if (andereRunde.Figuranten.Contains(gewaehlterFigurant))
                            {
                                figurantFoundCounter++;
                            }
                        }

                        int anzahlRundenMax;
                        if (gewaehlterFigurant.IsMitHund == false) //Figuranten ohne Hund bleiben länger draussen
                        {
                            anzahlRundenMax = calcSetting.AnzahlRundenDraussenKeinHF;
                        }
                        else
                        {
                            anzahlRundenMax = calcSetting.AnzahlRundenDraussen;
                        }

                        if (figurantFoundCounter >= anzahlRundenMax)
                        {
                            //schon zu oft figurant
                            figurantenGeloeschtWegenMaxAnzahl.Add(gewaehlterFigurant);
                            figuranten.RemoveAll(x => x == gewaehlterFigurant);
                        }
                    }

                    //Figurant laden für jeweilige Position
                    if (figuranten.Count == 0)
                    {
                        //nochmals jemanden nehmen der eigentlich die max anzahl erreicht hat
                        figuranten.AddRange(figurantenGeloeschtWegenMaxAnzahl);
                        runde.Infos.Add("Figurant auf Platz " + figurantPlatz.ToString() + " liegt öfters als geplant");

                        if (figuranten.Count == 0)
                        {
                            //in dem Fall wurde zuviel rausgelöscht, der nächste HF muss nochmals liegen
                            if (uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault() != null)
                            {
                                figuranten.Add(uebungsplan.Where(y => y.Order == runde.Order + 1).FirstOrDefault().Hundefuehrer);
                                runde.Infos.Add("Figurant kommt nachher sofort als HF dran");
                            }
                            else
                            {
                                //wenn es leider immer noch nicht ausreicht, den vorherigen HF  
                                if (uebungsplan.Where(y => y.Order == runde.Order - 1).FirstOrDefault() != null)
                                {
                                    figuranten.Add(uebungsplan.Where(y => y.Order == runde.Order - 1).FirstOrDefault().Hundefuehrer);
                                    runde.Infos.Add("Figurant war vorher als HF dran");
                                }
                            }
                        }
                    }
                    if (figuranten.Count > 0)
                    {
                        runde.Figuranten.Add(figuranten.OrderBy(x => x.IsMitHund).FirstOrDefault());
                    }
                    else
                    {
                        //error handling: Berechnung nicht möglich
                        runde.Infos.Add("Berechnung nicht möglich. Nicht genügend Figuranten in Figurant-Platz " + figurantPlatz.ToString());
                    }

                }
            }







        }
    }
}
