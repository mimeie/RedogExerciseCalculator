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

        private void Redog_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Initialize_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;            
            Excel.Worksheet konfiguration;

            
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



            ExerciseCalculator ec = new ExerciseCalculator();

            Excel.Window window = e.Control.Context;            
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);

            foreach (string dog in ec.Execute())
            {
                
                Excel.Range firstRow = activeWorksheet.get_Range("A1");
                firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
                newFirstRow.Value2 = dog;

            }

            int i = 0;
            foreach (string dog in ec.Execute())
            {
                i++;
                Excel.Range firstRow2 = activeWorksheet.get_Range("C" + i.ToString());
                firstRow2.Value2 = dog;                             
            }
        }
    }
}
