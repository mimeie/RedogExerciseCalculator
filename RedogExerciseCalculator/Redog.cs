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

            range  = konfiguration.get_Range("A1");
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
            range.Value2 = "1";
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = "(nicht implementiert) 0: Keine Mitte, 1: Mitte-Läufe am Stück, 2: Mitte abwechselnd";


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
            i = 11;
            
            addTeilnehmer("Joli", "x", "x", i++);
            addTeilnehmer("Sarah", "x", "x", i++);
            addTeilnehmer("Tom", "x", "", i++);
            addTeilnehmer("Alain", "x", "", i++);
            addTeilnehmer("Clemens", "x", "", i++);
            addTeilnehmer("Masha", "x", "", i++);
            addTeilnehmer("Alain", "x", "", i++);
            addTeilnehmer("Bettina", "x", "", i++);
            addTeilnehmer("Pascal", "x", "", i++);
            addTeilnehmer("Tom", "", "", i++);
            addTeilnehmer("Stefan", "", "", i++);

            //breite setzen, funktioniert noch nicht
            double width;
            width = ((Excel.Range)konfiguration.Cells[1, 1]).Width;
            ((Excel.Range)konfiguration.Cells[1, 1]).ColumnWidth = width;
            width = ((Excel.Range)konfiguration.Cells[1, 3]).Width;
            ((Excel.Range)konfiguration.Cells[1, 3]).ColumnWidth = width + 15;
        }


        private void addTeilnehmer(string name, string isMitHund, string isMitte, int i)
        {
            Excel.Range range;

            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = name;
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = isMitHund;
            range = konfiguration.get_Range("C" + i.ToString());
            range.Value2 = isMitte;
        }

        //private void DisplayActiveSheetName()
        //{
        //    Excel.Worksheet worksheet1 = (Excel.Worksheet)this.ActiveSheet;
        //    MessageBox.Show("The name of the active sheet is: " +
        //        worksheet1.Name);
        //}


        //        xlWorkSheet = (Worksheet) xlWorkBook.Worksheets["Dashboard"];
        //        range = xlWorkSheet.Cells[1, 1];
        //range.Value2 = "Test";

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
                MessageBox.Show("Übung kann nicht berechnet werden. Konfiguration nicht gefunden.","Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    range = konfiguration.get_Range("A" + i.ToString());
                    name = range.Value2;
                    range = konfiguration.get_Range("B" + i.ToString());
                    isMitHund = range.Value2;
                    range = konfiguration.get_Range("C" + i.ToString());
                    isMitte=  range.Value2;

                    if (name != null)
                    { 
                        ec.addTeilnehmer(name, isMitHund, isMitte);
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
                //worksheet neu eröffnen
                logger.Warn(ex, $"Problem beim Berechnen, vermutlich existiert Sheet noch nicht.");
                uebung = (Excel.Worksheet)window.Application.Worksheets.Add();
                uebung.Name = "Übungsplan";
            }

            ec.Execute();

            
            //foreach (string dog in ec.Execute())
            //{
            //    i++;
            //    Excel.Range firstRow2 = activeWorksheet.get_Range("C" + i.ToString());
            //    firstRow2.Value2 = dog;                             
            //}



            //Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);

            //foreach (string dog in ec.Execute())
            //{

            //    Excel.Range firstRow = activeWorksheet.get_Range("A1");
            //    firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            //    Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            //    newFirstRow.Value2 = dog;

            //}

            //int i = 0;
            //foreach (string dog in ec.Execute())
            //{
            //    i++;
            //    Excel.Range firstRow2 = activeWorksheet.get_Range("C" + i.ToString());
            //    firstRow2.Value2 = dog;                             
            //}
        }

        private void Info_Click(object sender, RibbonControlEventArgs e)
        {
            Info info = new Info();
            info.ShowDialog();

        }
    }
}
