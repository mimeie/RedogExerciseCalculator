using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using RedogExerciseBusinessLogic;

namespace RedogExerciseCalculatorTest
{
    [TestClass]
    public class MainTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            ExerciseCalculator ec = new ExerciseCalculator();
            ec.Execute();

            foreach (Uebungsrunde hf in ec.uebungsplan)
            {
                Console.WriteLine(hf.Hundefuehrer.Name);
            }
        }
    }
}
