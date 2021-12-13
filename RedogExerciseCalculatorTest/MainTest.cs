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
            foreach (string dog in ec.Execute())
            {
                Console.WriteLine(dog);
            }
        }
    }
}
