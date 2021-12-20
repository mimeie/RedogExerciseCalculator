using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using System.Collections.Generic;
using System.Linq;

using RedogExerciseBusinessLogic;

using NLog;
using NLog.Config;
using NLog.Targets.ElasticSearch;

namespace RedogExerciseCalculatorTest
{
    [TestClass]
    public class MainTest
    {
      

        [TestMethod]
        public void TestMethod1()
        {

            // enable internal logging to the console
            NLog.Common.InternalLogger.LogToConsole = true;

            //// enable internal logging to a file
            NLog.Common.InternalLogger.LogFile = "log.txt";

            //// enable internal logging to a custom TextWriter
           // NLog.Common.InternalLogger.LogWriter = new StringWriter(); //e.g. TextWriter writer = File.CreateText("C:\\perl.txt")

            //// set internal log level
            NLog.Common.InternalLogger.LogLevel = LogLevel.Trace;

            var config = new LoggingConfiguration();
            ElasticSearchTarget elastictarget = new ElasticSearchTarget
            {
                Name = "elastic",
                Uri = "http://jhistorian.prod.j1:9200/",  //Uri = "http://192.168.2.41:32120", 
                Index = "RedogExerciseCalculator-Test-${level}-${date:format=yyyy-MM-dd}",
                //Index = "historianWriter-${level}-${date:format=yyyy-MM-dd}",
                //Layout = "${logger} | ${threadid} | ${message}",
                //Layout = "${longdate}|${event-properties:item=EventId_Id}|${threadid}|${uppercase:${level}}|${logger}|${hostname}|${message} ${exception:format=tostring}",
                Layout = "${message}",
                IncludeAllProperties = true,
            };
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, elastictarget, "*");
            LogManager.Configuration = config;

            Logger log = LogManager.GetCurrentClassLogger();


            log.Info("Funktionstest gestartet");

          

            ExerciseCalculator ec = new ExerciseCalculator();

            ec.calcSetting.AnzahlFiguranten = 3;
            ec.calcSetting.AnzahlRundenDraussen = 3;
            ec.calcSetting.AnzahlRundenDraussenKeinHF = 4;

            ec.addTeilnehmer("Michael", true,false);
            ec.addTeilnehmer("Sarah", true, true);
            ec.addTeilnehmer("Tom", false, false);
            ec.addTeilnehmer("Alain", true, false);
            ec.addTeilnehmer("Clemens", true, false);
            ec.addTeilnehmer("Joli", true, true);
            ec.addTeilnehmer("Masha", true, false);
            ec.addTeilnehmer("Pascal", true, false);
            ec.addTeilnehmer("Bettina", true, false);
            ec.addTeilnehmer("Duri", true, false);
            ec.addTeilnehmer("Stefan", false, false);

            ec.Execute();


          
            foreach (Uebungsrunde runde in ec.uebungsplan.ToList().OrderBy(x => x.Order))
            {
                List<string> figurantenNamen = new List<string>();           
                foreach (Teilnehmer figurant in runde.Figuranten.OrderBy(x => x.FigurantenPlatz))
                {
                    figurantenNamen.Add(figurant.FigurantenPlatz.ToString() + ") " + figurant.Name);
                 
                }

                log.Info("Suche {0}, HF: {1}, Mitte: {2}, Figuranten: {3}, Info:{4}", runde.Order, runde.Hundefuehrer.Name,runde.Mitte.Name, string.Join(", ", figurantenNamen.ToList()),runde.Info);
            }
        }
    }
}
