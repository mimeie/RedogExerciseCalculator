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

            ec.calcSetting.AnzahlFiguranten = 2;
            ec.calcSetting.AnzahlRundenDraussen = 2;
            ec.calcSetting.AnzahlRundenDraussenKeinHF = 4;
            ec.calcSetting.KonfigurationMitte = KonfigurationMitte.MitteAbwechselnd;

            ec.addTeilnehmer("Michael", true,false,false);
            ec.addTeilnehmer("Sarah", true, true, false);
            ec.addTeilnehmer("Tom", false, false, false);
            ec.addTeilnehmer("Alain", true, false, false);
            ec.addTeilnehmer("Clemens", true, false, false);
            ec.addTeilnehmer("Joli", true, true, false);
            ec.addTeilnehmer("Masha", true, false, true);
            ec.addTeilnehmer("Pascal", true, false, false);
            ec.addTeilnehmer("Bettina", true, false, true);
            ec.addTeilnehmer("Duri", true, false, false);
            ec.addTeilnehmer("Stefan", false, false, false);

            ec.Execute();


          
            foreach (Uebungsrunde runde in ec.uebungsplan.ToList().OrderBy(x => x.Order))
            {
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

                log.Info("Suche {0}, HF: {1}, Mitte: {2}, Figuranten: {3}, Info:{4}", runde.Order, runde.Hundefuehrer.Name,runde.Mitte.Name, string.Join(", ", figurantenNamen.ToList()), string.Join(", ", infoTexte.ToList()));
            }
        }
    }
}
