using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RedogExerciseBusinessLogic;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using NLog;


namespace RedogExerciseCalculator
{
    public partial class Redog
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        Excel.Window window;
        Excel.Worksheet konfiguration;
        Excel.Worksheet uebung;
        ExerciseCalculator ec;

        private void Redog_Load(object sender, RibbonUIEventArgs e)
        {
            ec = new ExerciseCalculator();
        }

        private void Initialize_Click(object sender, RibbonControlEventArgs e)
        {
            window = e.Control.Context;



            //Konfig sheet anlegen
            try
            {
                konfiguration = (Excel.Worksheet)window.Application.Worksheets["Konfiguration"];
            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim Initalisieren, vermutlich existiert Sheet noch nicht.");
                konfiguration = (Excel.Worksheet)window.Application.Worksheets.Add();
                konfiguration.Name = "Konfiguration";
            }

            //überflüssige Tabelle löschen
            try
            {
                var tabelle1 = (Excel.Worksheet)window.Application.Worksheets["Tabelle1"].Delete();
            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim Initalisieren, Tabelle 1 schon gelöscht");
            }



            //Standardinformationen abfüllen
            Excel.Range range;
            int i = 1;

            range = konfiguration.get_Range("A1");
            range.Value2 = "Basis-Einstellungen";
            range.Font.Bold = true;
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "Wert";
            range.Font.Bold = true;
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "Hinweis/Bemerkung";
            range.Font.Bold = true;

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Anzahl Figuranten";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "2";
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "Eingesetzte Figuranten bei einer Suche";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Anzahl Runden draussen als Figurant";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "2";
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "Wird bei Figurantenmangel ignoriert";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Anzahl Runden draussen als Figurant (kein HF)";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "4";
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "Wird bei Figurantenmangel ignoriert";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Konfiguration Mitte";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "2";
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "0: Keine Mitte, 1: Mitte-Läufe am Stück, 2: Mitte abwechselnd";


            //Teilnehmer/Hundeführer abfüllen
            i = 10;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Teilnehmer";
            range.Font.Bold = true;
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "mit Hund";
            range.Font.Bold = true;
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "Mitte";
            range.Font.Bold = true;
            range = konfiguration.get_Range("D" + i.ToString());
            range.Value2 = "Spezialprogramm Anfang";
            range.Font.Bold = true;
            i = 11;

            addTeilnehmer("Joli", "x", "x", "", i++);
            addTeilnehmer("Sarah", "x", "x", "", i++);
            addTeilnehmer("Alain", "x", "", "", i++);
            addTeilnehmer("Clemens", "x", "", "", i++);
            addTeilnehmer("Masha", "x", "", "x", i++);
            addTeilnehmer("Michael", "x", "", "", i++);
            addTeilnehmer("Alain", "x", "", "", i++);
            addTeilnehmer("Bettina", "x", "", "x", i++);
            addTeilnehmer("Pascal", "x", "", "", i++);
            addTeilnehmer("Tom", "", "", "", i++);
            addTeilnehmer("Stefan", "", "", "", i++);

            //breite setzen, funktioniert nur so halb
            try
            {
                double width;
                width = ((Excel.Range)konfiguration.Cells[1, 1]).Width;
                ((Excel.Range)konfiguration.Cells[1, 1]).ColumnWidth = width;
                width = ((Excel.Range)konfiguration.Cells[1, 3]).Width;
                ((Excel.Range)konfiguration.Cells[1, 3]).ColumnWidth = width + 15;
            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Fehler beim Breite setzen");
            }

        }


        private void addTeilnehmer(string name, string isMitHund, string isMitte,string isSpezialprogrammAnfang, int i)
        {
            Excel.Range range;

            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = name;
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = isMitHund;
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = isMitte;
            range = konfiguration.get_Range("D" + i.ToString());
            range.Value2 = isSpezialprogrammAnfang;
        }



        private void getConfigToBusinessLogic()
        {
            try
            {
                konfiguration = (Excel.Worksheet)window.Application.Worksheets["Konfiguration"];
            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim Initalisieren, vermutlich existiert Sheet noch nicht.");
                MessageBox.Show("Übung kann nicht berechnet werden. Konfiguration nicht gefunden.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //basis Settings laden

            Excel.Range range;

            try
            {
                //basis konfig laden
                range = konfiguration.get_Range("B2");
                ec.calcSetting.AnzahlFiguranten = (int)range.Value2;

                range = konfiguration.get_Range("B3");
                ec.calcSetting.AnzahlRundenDraussen = (int)range.Value2;

                range = konfiguration.get_Range("B4");
                ec.calcSetting.AnzahlRundenDraussenKeinHF = (int)range.Value2;

                range = konfiguration.get_Range("B5");
                ec.calcSetting.KonfigurationMitte = (KonfigurationMitte)range.Value2;

                //teilnehmer ab Zeile 11 durchladen, einfach fix für 100 HF
                for (int i = 11; i < 100; i++)
                {
                    string name;
                    string isMitHund;
                    string isMitte;
                    string isSpezialprogrammAnfang;

                    range = konfiguration.get_Range("A" + i.ToString());
                    name = range.Value2;
                    range = konfiguration.get_Range("B" + i.ToString());
                    isMitHund = range.Value2;
                    range = konfiguration.get_Range("C" + i.ToString());
                    isMitte = range.Value2;
                    range = konfiguration.get_Range("D" + i.ToString());
                    isSpezialprogrammAnfang = range.Value2;

                    if (name != null)
                    {
                        ec.addTeilnehmer(name, isMitHund, isMitte, isSpezialprogrammAnfang);
                    }
                }


            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim laden der Konfiguration. Vermutlich ist die Konfiguration nicht lesbar.");
                MessageBox.Show("Problem beim laden der Konfiguration. Vermutlich ist die Konfiguration nicht lesbar.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }







        }


        private void Calculate_Click(object sender, RibbonControlEventArgs e)
        {

            ec.Initialize();

            window = e.Control.Context;


            //Konfiguration auslesen
            getConfigToBusinessLogic();

            //überflüssige Tabelle löschen

            try
            {
                var uebungsplan = (Excel.Worksheet)window.Application.Worksheets["Übungsplan"].Delete();
            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim Berechnen, übungsplan schon gelöscht");
            }

            //übungs sheet anlegen
            try
            {
                uebung = (Excel.Worksheet)window.Application.Worksheets["Übungsplan"];





            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen, weiss nicht wie es besser geht
                logger.Warn(ex, $"Problem beim Berechnen, vermutlich existiert Sheet noch nicht.");
                uebung = (Excel.Worksheet)window.Application.Worksheets.Add();
                uebung.Name = "Übungsplan";

                ec.Execute();

                Excel.Range range;
                int i = 1;

                range = uebung.get_Range("A" + i.ToString());
                range.Value2 = "Übungsplan";
                range.Font.Bold = true;
                range.Font.Size = 14;
                i = 3;

                range = uebung.get_Range("A" + i.ToString());
                range.Value2 = "Runde";
                range.Font.Bold = true;
                range = uebung.get_Range("B" + i.ToString());
                range.Value2 = "Zeit";
                range.Font.Bold = true;
                range = uebung.get_Range("C" + i.ToString());
                range.Value2 = "Hundeführer";
                range.Font.Bold = true;
                range = uebung.get_Range("D" + i.ToString());
                range.Value2 = "Mitte";
                range.Font.Bold = true;
                range = uebung.get_Range("E" + i.ToString());
                range.Value2 = "Figuranten";
                range.Font.Bold = true;
                range = uebung.get_Range("F" + i.ToString());
                range.Value2 = "Infos";
                range.Font.Bold = true;




                foreach (Uebungsrunde runde in ec.uebungsplan.ToList().OrderBy(x => x.Order))
                {
                    i++;
                    List<string> figurantenNamen = new List<string>();
                    foreach (Teilnehmer figurant in runde.Figuranten.OrderBy(x => x.FigurantenPlatz))
                    {
                        figurantenNamen.Add(figurant.FigurantenPlatz.ToString() + ") " + figurant.Name);
                    }

                    List<string> infoTexte = new List<string>();
                    foreach (string info in runde.Infos)
                    {
                        infoTexte.Add(info);
                    }

                    range = uebung.get_Range("A" + i.ToString());
                    range.Value2 = runde.Order.ToString();
                    range = uebung.get_Range("B" + i.ToString());
                    range.Value2 = "";
                    range = uebung.get_Range("C" + i.ToString());
                    range.Value2 = runde.Hundefuehrer.Name;
                    range = uebung.get_Range("D" + i.ToString());
                    if (runde.Mitte != null)
                    {
                        range.Value2 = runde.Mitte.Name;
                    }
                    range = uebung.get_Range("E" + i.ToString());
                    range.Value2 = string.Join(", ", figurantenNamen.ToList());
                    range = uebung.get_Range("F" + i.ToString());
                    range.Value2 = string.Join(", ", infoTexte.ToList());


                    //log.Info("Suche {0}, HF: {1}, Mitte: {2}, Figuranten: {3}, Info:{4}", runde.Order, runde.Hundefuehrer.Name, runde.Mitte.Name, string.Join(", ", figurantenNamen.ToList()), runde.Info);
                }


            }

            //breite setzen, funktioniert nur so halb
            try
            {
                double width;
                width = ((Excel.Range)uebung.Cells[1, 5]).Width;
                ((Excel.Range)uebung.Cells[1, 5]).ColumnWidth = width;

            }
            catch (Exception ex)
            {
                //worksheet neu eröffnen
                logger.Warn(ex, $"Fehler beim Breite setzen");
            }

        }

        private void Info_Click(object sender, RibbonControlEventArgs e)
        {
            Info info = new Info();
            info.ShowDialog();

        }
    }
}
