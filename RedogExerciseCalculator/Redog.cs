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

        private void Redog_Load(object sender, RibbonUIEventArgs e)
        {

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
            range.Value2 = "Basis-Einstellungen:";
            range.Font.Bold = true;

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Anzahl Figuranten";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "1";


            //Teilnehmer/Hundeführer abfüllen
            i = 10;           
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Teilnehmer";
            range.Font.Bold = true;
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "mit Hund";
            range.Font.Bold = true;


            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Michael";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "x";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Alain";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "x";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Tom";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Sarah";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "x";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Clemens";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "x";

            i++;
            range = konfiguration.get_Range("A" + i.ToString());
            range.Value2 = "Duri";
            range = konfiguration.get_Range("B" + i.ToString());
            range.Value2 = "x";
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


        private void Calculate_Click(object sender, RibbonControlEventArgs e)
        {



            
            window = e.Control.Context;

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

            ExerciseCalculator ec = new ExerciseCalculator();
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
    }
}
