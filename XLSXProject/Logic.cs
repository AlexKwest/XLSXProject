using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using XLSXProject.model;
using OfficeOpenXml.FormulaParsing;

namespace XLSXProject
{

    public class Logic
    {
        ExcelPackage excelIn;
        ExcelPackage excelOut;
        ExcelWorksheet worksheet;

        ExcelWorksheet motivashion;
        ExcelWorksheet newSheet;
        List<OperatorModel> operatorModels;

       
        public Logic(string FilePathIn, string FilePathOut)
        {
            excelIn = new ExcelPackage(new FileInfo(FilePathIn));
            excelOut = new ExcelPackage(new FileInfo(FilePathOut));
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
            worksheet = excelIn.Workbook.Worksheets[numberSheet];

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
            /* TODO: На время отладки    
                Console.WriteLine("Введите Оклад сотрудника:" + result.Name);
                result.Oklad = Convert.ToInt32(Console.ReadLine()); 
            */
            result.Oklad = 6250;
                return result;
        }
        private bool GetValueDay(int row)
        {
            try
            {
                if (worksheet.Cells[row, 1].Value.ToString() == "потеряшки")
                    return true;
            }
            catch { }
            return Convert.ToInt32(worksheet.Cells[row, 62].Value) != 0 ? true : false;
        }

        private bool GetValueOperator(int row)
        {
            return worksheet.Cells[row, 1].Value != null ? true : false;
        }

        //TODO: Респечатать результат с 1-15 число
        public void PrintResult(string sheetName, EnumResult.PrintFile typePrint)
        {
            switch (typePrint)
            {
                case EnumResult.PrintFile.FirsMonth:  PrintFirstPathMonth(sheetName); break;
                case EnumResult.PrintFile.TwoMonth:   PrintTwoPathMonth(sheetName); break;
            }

            excelOut.Save();
        }

        private void GreateNewSheet(string sheetName)
        {
            try
            {
                newSheet = excelOut.Workbook.Worksheets.Add(sheetName);
            }
            catch
            {
                newSheet = excelOut.Workbook.Worksheets[sheetName];
            }
        }

        private void GreateMotivationSheet()
        {
            try
            {
                motivashion = excelOut.Workbook.Worksheets.Add("Мотивация");
                //
                //worksheet.Cells[FromRow, FromColumn, ToRow, ToColumn].Merge = true;
                motivashion.Cells[1, 1, 1, 3].Merge = true;
                motivashion.Cells[1, 1, 7, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                motivashion.Cells[1, 1, 1, 3].Style.Font.Color.SetColor(Color.DarkRed);
                motivashion.Cells[1, 1, 7, 3].Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.DarkBlue);
                motivashion.SetValue(1, 1, 6250);
                motivashion.Cells[2, 1, 2, 3].Merge = true;
                motivashion.SetValue(2, 1, "С 1-15 оклад");
                motivashion.SetValue(3, 1, "от"); motivashion.SetValue(3, 2, "до"); motivashion.SetValue(3, 3, "сумма");

                motivashion.SetValue(4, 1, 1); motivashion.SetValue(4, 2, 4); motivashion.SetValue(4, 3, 200);
                motivashion.SetValue(5, 1, 5); motivashion.SetValue(5, 2, 9); motivashion.SetValue(5, 3, 400);
                motivashion.SetValue(6, 1, 10); motivashion.SetValue(6, 2, 14); motivashion.SetValue(6, 3, 600);
                motivashion.SetValue(7, 1, 15); motivashion.SetValue(7, 3, 1000);

                motivashion.Cells[1, 5, 1, 7].Merge = true;
                motivashion.Cells[1, 5, 1, 7].Style.Font.Color.SetColor(Color.DarkRed);
                motivashion.Cells[1, 5, 7, 7].Style.Border.BorderAround(ExcelBorderStyle.DashDot,Color.DarkBlue);
                motivashion.SetValue(1, 5, 10000);
                motivashion.Cells[2, 5, 2, 7].Merge = true;
                motivashion.SetValue(2, 5, "С 1-15 оклад");
                motivashion.SetValue(3, 5, "от"); motivashion.SetValue(3, 6, "до"); motivashion.SetValue(3, 7, "сумма");

                motivashion.SetValue(4, 5, 1); motivashion.SetValue(4, 6, 4); motivashion.SetValue(4, 7, 50);
                motivashion.SetValue(5, 5, 11); motivashion.SetValue(5, 7, 100);
            }
            catch
            {
                motivashion = excelOut.Workbook.Worksheets["Мотивация"];
            }
        }

        private void PrintTwoPathMonth(string sheetName)
        {
            GreateMotivationSheet();
        }

        public int GetPointPoteriashka()
        {
            return OperatorModel.Poteriashki;
            //return    (from t in operatorModels
            //                where t.Name == "потеряшки"
            //                select t.Proideno15).Sum();
        }

        

        private void PrintFirstPathMonth(string sheetName)
        {
          

            GreateNewSheet(sheetName);
            GreateMotivationSheet();
            var row = 2;
            var col = 1;
            //newSheet.Column(1).Width = 30;
            newSheet.SetValue(1, 1, "Оператор");
            newSheet.SetValue(1, 2, "День");
            newSheet.SetValue(1, 3, "Пройдено");
            newSheet.SetValue(1, 4, "Оклад");
            newSheet.SetValue(1, 5, "Бонус");
            newSheet.SetValue(1, 6, "Оклад к выплате");
            newSheet.SetValue(1, 7, "Бонусы 1 - 15");
            newSheet.SetValue(1, 8, "Сумма");

            foreach (var cell in operatorModels)
            {
                if (cell.Name != "потеряшки")
                {
                    newSheet.SetValue(row, col++, cell.Name);
                    newSheet.SetValue(row, col++, cell.Days15);
                    newSheet.SetValue(row, col++, cell.Proideno15);
                    newSheet.SetValue(row, col++, cell.Oklad);

                    newSheet.Cells[row, col++].Formula = "IF(" + newSheet.Cells[row, 4].Address + "=" + motivashion.Name + "!" + motivashion.Cells[1, 1].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + ">=" + motivashion.Name + "!" + motivashion.Cells[4, 1].Address + "," + newSheet.Cells[row, 3].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[4, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[4, 3].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[5, 1].Address + ", " + newSheet.Cells[row, 3].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[5, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[5, 3].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[6, 1].Address + ", " + newSheet.Cells[row, 3].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[6, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[6, 3].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[7, 1].Address + ")," + motivashion.Name + "!" + motivashion.Cells[7, 3].Address + ",0" + ")))),"
                                                       + "IF(" + newSheet.Cells[row, 4].Address + "=" + motivashion.Name + "!" + motivashion.Cells[1, 5].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[4, 5].Address + ", " + newSheet.Cells[row, 3].Address + " <= " + motivashion.Name + "!" + motivashion.Cells[4, 6].Address + "), " + motivashion.Name + "!" + motivashion.Cells[4, 7].Address
                                                                                                                                                     + ", IF(AND(" + newSheet.Cells[row, 3].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[5, 5].Address + ")," + motivashion.Name + "!" + motivashion.Cells[5, 7].Address + ",0" + ")),0))";
                    // wr.SetValue(row, col++, cell.Bonus);

                    newSheet.Cells[row, col++].Formula = "=(" + newSheet.Cells[row, 4].Address + "/" + 25 + ")*" + newSheet.Cells[row, 2].Address;
                    //wr.SetValue(row, col++, cell.OkladinPay);

                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 3].Address + "*" + newSheet.Cells[row, 5].Address;
                    //wr.SetValue(row, col++, cell.BonusDyas15);

                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 6].Address + "+" + newSheet.Cells[row, 7].Address;
                    //wr.SetValue(row, col++, cell.Summa);
                    col = 1;
                    row++;
                }
            }
            newSheet.SetValue(row + 3, 1, "Распределите очки Потеряшки: " + GetPointPoteriashka());
            StyleTable();
        }

        private void StyleTable()
        {
            newSheet.Column(1).AutoFit(10);
            newSheet.Column(2).AutoFit(10);
            newSheet.Column(3).AutoFit(10);
            newSheet.Column(4).AutoFit(10);
            newSheet.Column(5).AutoFit(10);
            newSheet.Column(6).AutoFit(10);
            newSheet.Column(7).AutoFit(10);
            newSheet.Column(8).AutoFit(10);

            var modelRows = operatorModels.Count(x => x.Name != "потеряшки")+ 1;
            string modelRange = "A1:H" + modelRows.ToString();
            var modelTable = newSheet.Cells[modelRange];

            // Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //Style Head table
            modelTable = newSheet.Cells["A1:H1"];
            modelTable.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            modelTable.Style.Font.Color.SetColor(Color.DarkGreen);
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }
    }

}

