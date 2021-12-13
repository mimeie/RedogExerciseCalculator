using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RedogExerciseBusinessLogic;


namespace RedogExerciseCalculator
{
    public partial class Redog
    {
        private void Redog_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Initialize_Click(object sender, RibbonControlEventArgs e)
        {
            
           
        }

        private void Calculate_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(Class1.getTestresult());
        }
    }
}
