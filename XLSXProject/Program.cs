using System;
using System.Collections.Generic;
using XLSXProject.model;

namespace XLSXProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Logic logic = new Logic(@"c:\Project\XLSX\XLSXProject\XLSXProject\data\Operators.xlsx");


            List<OperatorModel> operatorModels = logic.SetOperatorList();
           // operatorModels = operatorModels.Distinct(new MyFile()).ToList();

            Console.SetWindowSize(160, Console.WindowHeight * 2);
            foreach (var result in operatorModels)
            {
                Console.WriteLine(result.Show());
            }

            logic.PrintResult(@"c:\Project\XLSX\XLSXProject\XLSXProject\data\Result.xlsx");
            Console.WriteLine("Hello World!");
            Console.WriteLine("Press any key...");
            Console.ReadKey();

        }
    }
}
