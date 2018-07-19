using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;
using XLSXProject.model;

namespace XLSXProject
{

    public class Logic
    {
        ExcelPackage excel;
        ExcelWorksheet worksheet;
        List<OperatorModel> operatorModels;

        public Logic(string FilePath)
        {
            FileInfo existingFile = new FileInfo(FilePath);
            excel = new ExcelPackage(existingFile);
            operatorModels = new List<OperatorModel>();
        }

        public List<OperatorModel> SetOperatorList()
        {
            GetListSheet(0);
            GetListSheet(1);
            return operatorModels;
        }

        private void GetListSheet(int numberSheet)
        {
            worksheet = excel.Workbook.Worksheets[numberSheet];

            for (var row = 3; row < 100; row++)
            {
                if (GetValueOperator(row))
                {
                    if (GetValueDay(row))
                        operatorModels.Add(SetOperator(row));
                }
                else break;
            }
        }

        private OperatorModel SetOperator(int row)
        {
                var result = new OperatorModel();
                result.Name = worksheet.Cells[row, 1].Value.ToString();
                result.Days15 = Convert.ToInt32(worksheet.Cells[row, 62].Value);
                result.Proideno15 = Convert.ToInt32(worksheet.Cells[row, 64].Value);
                Console.WriteLine("Введите Оклад сотрудника:" + result.Name);
                // TODO: На время отладки result.Oklad = Convert.ToInt32(Console.ReadLine()); 
                result.Oklad = 6250;
                return result;
        }
        private bool GetValueDay(int row)
        {
            return Convert.ToInt32(worksheet.Cells[row, 62].Value) != 0 ? true : false;
        }

        private bool GetValueOperator(int row)
        {
            return worksheet.Cells[row, 1].Value != null ? true : false;
        }

        public void PrintResult(string fileName)
        {
            var excelSave = new ExcelPackage();
            var wr = excelSave.Workbook.Worksheets.Add("Result");
            var row = 2;
            var col = 1;

            wr.SetValue(1, 1, "Оператор");
            wr.SetValue(1, 2, "День");
            wr.SetValue(1, 3, "Пройдено");
            wr.SetValue(1, 4, "Оклад");
            wr.SetValue(1, 5, "Бонус");
            wr.SetValue(1, 6, "Оклад к выплате");
            wr.SetValue(1, 7, "Бонусы 1 - 15");
            wr.SetValue(1, 8, "Сумма");

            foreach (var cell in operatorModels)
            {
                wr.SetValue(row, col++, cell.Name);
                wr.SetValue(row, col++, cell.Days15);
                wr.SetValue(row, col++, cell.Proideno15);
                wr.SetValue(row, col++, cell.Oklad);
                wr.SetValue(row, col++, cell.Bonus);

                wr.Cells[row, col++].Formula = "=(" + wr.Cells[row,4].Address + "/" + 25 + ")*" + wr.Cells[row,2].Address;
                //wr.SetValue(row, col++, cell.OkladinPay);

                wr.Cells[row, col++].Formula = "=" + wr.Cells[row, 3].Address + "*" + wr.Cells[row, 5].Address ;
                //wr.SetValue(row, col++, cell.BonusDyas15);

                wr.Cells[row, col++].Formula = "=" + wr.Cells[row, 6].Address + "+" + wr.Cells[row, 7].Address;
                //wr.SetValue(row, col++, cell.Summa);
                col = 1;
                row++;
            }

            excelSave.SaveAs(new FileInfo(fileName));

            //TODO: Респечатать результат с 1-15 число
        }
    }

}


//// output the formula in row 5
//Console.WriteLine("\tCell({0},{1}).Formula={2}", 3, 5, worksheet.Cells[3, 5].Formula);                
//            Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells[3, 5].FormulaR1C1);

//            // output the formula in row 5
//            Console.WriteLine("\tCell({0},{1}).Formula={2}", 5, 3, worksheet.Cells[5, 3].Formula);
//            Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells[5, 3].FormulaR1C1);

