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


namespace RedogExerciseCalculator
{
    public partial class Redog
    {
        private void Redog_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Initialize_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(Class1.getTestresult());

        }

        //private void DisplayActiveSheetName()
        //{
        //    Excel.Worksheet worksheet1 = (Excel.Worksheet)this.ActiveSheet;
        //    MessageBox.Show("The name of the active sheet is: " +
        //        worksheet1.Name);
        //}

        private void Calculate_Click(object sender, RibbonControlEventArgs e)
        {
            ExerciseCalculator ec = new ExerciseCalculator();
           
            foreach (string dog in ec.Execute())
            {
                Excel.Window window = e.Control.Context;              
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
                Excel.Range firstRow = activeWorksheet.get_Range("A1");
                firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
                newFirstRow.Value2 = dog;

            }
        }
    }
}
