﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLSXProject.model;

namespace XLSXProject
{
    
    class Program
    {
        static string demoPathIn = Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName,"data","Operators.xlsx");
        static string demoPathOut = Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).FullName, "data", "Result.xlsx");
        public const string poteriashka = "потеряшки";
        static void Main(string[] args)
        {
            Console.SetWindowSize(160, Console.WindowHeight * 2);

            Logic logic = new Logic(demoPathIn, demoPathOut);
            List<OperatorModel> operatorModels = logic.SetOperatorList();

            foreach (var result in operatorModels)
            {
                Console.WriteLine(result.Show());
            }

            logic.PrintResult("1-15 Операторы", EnumResult.PrintFile.FirsMonth);
            logic.PrintResult("16-31 Операторы", EnumResult.PrintFile.TwoMonth);
           
            Console.WriteLine("Press any key...");

            Console.WriteLine("У потеряшки было: " + logic.GetPointPoteriashka(EnumResult.PrintFile.FirsMonth));

            Console.ReadKey();

        }

    }
}
