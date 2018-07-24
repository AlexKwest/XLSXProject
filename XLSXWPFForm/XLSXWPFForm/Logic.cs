using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using XLSXWPFForm.model;

namespace XLSXWPFForm
{
    public class Logic
    {
        ExcelPackage excelIn;
        ExcelPackage excelOut;
        ExcelWorksheet worksheet;

        ExcelWorksheet motivashion;
        ExcelWorksheet newSheet;
        List<OperatorModel> operatorModels;

        IEnumerable<OperatorModel> result;

        public EnumResult.InputOklad inputOklad;

        public Logic(string FilePathIn)
        {
            excelIn = new ExcelPackage(new FileInfo(FilePathIn));
            excelOut = new ExcelPackage();
            operatorModels = new List<OperatorModel>();
        }

      
        public List<OperatorModel> SetOperatorList(EnumResult.InputOklad inputOklad)
        {
            this.inputOklad = inputOklad;
            GetListSheet(1, 62);
            GetListSheet(2, 130);
            return operatorModels;
        }

        private void GetListSheet(int numberSheet, int proideno)
        {
            worksheet = excelIn.Workbook.Worksheets[numberSheet];

            for (var row = 3; row < 100; row++)
            {
                if (GetValueOperator(row))
                {
                    operatorModels.Add(SetOperator(row));
                }
                else break;
            }
        }

        private bool GetValueDay(int row, int proideno)
        {
            return Convert.ToInt32(worksheet.Cells[row, proideno].Value) != 0 ? true : false;
        }

        private bool GetValueOperator(int row)
        {
            return worksheet.Cells[row, 1].Value != null ? true : false;
        }

        private OperatorModel SetOperator(int row)
        {
            var result = new OperatorModel();
            result.Name = worksheet.Cells[row, 1].Value.ToString();
            result.Days15 = Convert.ToInt32(worksheet.Cells[row, 62].Value);
            result.Proideno15 = Convert.ToInt32(worksheet.Cells[row, 64].Value);

            if (inputOklad == EnumResult.InputOklad.Input)
            {
                var dialog = new InputBoxes(result.Name, this);
                dialog.ShowDialog();
                result.Oklad = Convert.ToInt32(dialog.OklasResult);
            }
            else if (inputOklad == EnumResult.InputOklad.Default)
            {
                // TODO: На время отладки    
                result.Oklad = 6250;
            }

            result.Days31 = Convert.ToInt32(worksheet.Cells[row, 130].Value);
            result.Proideno31 = Convert.ToInt32(worksheet.Cells[row, 132].Value);
            return result;
        }

        public void PrintResult(string sheetName, EnumResult.PrintFile typePrint, string demoFileOut)
        {
            switch (typePrint)
            {
                case EnumResult.PrintFile.FirsMonth: PrintFirstPathMonth(sheetName); break;
                case EnumResult.PrintFile.TwoMonth: PrintTwoPathMonth(sheetName); break;
            }

            excelOut.SaveAs(new FileInfo(demoFileOut));
        }

        private void GreateNewSheet(string sheetName)
        {
            try
            {
                newSheet = excelOut.Workbook.Worksheets.Add(sheetName);
            }
            catch
            {
                excelOut.Workbook.Worksheets.Delete(sheetName);
                GreateNewSheet(sheetName);
                // newSheet = excelOut.Workbook.Worksheets[sheetName];
            }
        }

        private void GreateMotivationSheet()
        {
            try
            {
                motivashion = excelOut.Workbook.Worksheets.Add("Мотивация");
                motivashion.Cells[1, 1, 30, 20].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                motivashion.Cells[1, 1, 1, 3].Merge = true;
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

                motivashion.Cells[10, 1, 10, 3].Merge = true;
                motivashion.Cells[10, 1, 10, 3].Style.Font.Color.SetColor(Color.DarkRed);
                motivashion.Cells[10, 1, 17, 3].Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.DarkBlue);
                motivashion.SetValue(10, 1, 6250);
                motivashion.Cells[11, 1, 11, 3].Merge = true;
                motivashion.SetValue(11, 1, "С 16-31 оклад");
                motivashion.SetValue(12, 1, "от"); motivashion.SetValue(12, 2, "до"); motivashion.SetValue(12, 3, "сумма");
                motivashion.SetValue(13, 1, 1); motivashion.SetValue(13, 2, 4); motivashion.SetValue(13, 3, 200);
                motivashion.SetValue(14, 1, 5); motivashion.SetValue(14, 2, 9); motivashion.SetValue(14, 3, 400);
                motivashion.SetValue(15, 1, 10); motivashion.SetValue(15, 2, 14); motivashion.SetValue(15, 3, 600);
                motivashion.SetValue(16, 1, 15); motivashion.SetValue(16, 2, 19); motivashion.SetValue(16, 3, 800);
                motivashion.SetValue(17, 1, 20); motivashion.SetValue(17, 3, 1000);


                motivashion.Cells[1, 5, 1, 7].Merge = true;
                motivashion.Cells[1, 5, 1, 7].Style.Font.Color.SetColor(Color.DarkRed);
                motivashion.Cells[1, 5, 7, 7].Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.DarkBlue);
                motivashion.SetValue(1, 5, 10000);
                motivashion.Cells[2, 5, 2, 7].Merge = true;
                motivashion.SetValue(2, 5, "С 1-15 оклад");
                motivashion.SetValue(3, 5, "от"); motivashion.SetValue(3, 6, "до"); motivashion.SetValue(3, 7, "сумма");
                motivashion.SetValue(4, 5, 1); motivashion.SetValue(4, 6, 4); motivashion.SetValue(4, 7, 50);
                motivashion.SetValue(5, 5, 11); motivashion.SetValue(5, 7, 100);

                motivashion.Cells[1, 9, 2, 11].Merge = true;
                motivashion.Cells[1, 9, 2, 11].Style.Font.Color.SetColor(Color.DarkRed);
                motivashion.Cells[1, 9, 7, 11].Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.DarkBlue);
                motivashion.SetValue(1, 9, "Премия");
                motivashion.SetValue(3, 9, "от"); motivashion.SetValue(3, 10, "до"); motivashion.SetValue(3, 11, "сумма");
                motivashion.SetValue(4, 9, 30); motivashion.SetValue(4, 10, 34); motivashion.SetValue(4, 11, 3000);
                motivashion.SetValue(5, 9, 35); motivashion.SetValue(5, 11, 5000);

            }
            catch
            {
                motivashion = excelOut.Workbook.Worksheets["Мотивация"];
            }
        }

        public int GetPointPoteriashka(EnumResult.PrintFile type)
        {
            switch (type)
            {
                case EnumResult.PrintFile.FirsMonth:
                    return (from t in operatorModels
                            where t.Name == MainWindow.poteriashka
                            select t.Proideno15).Sum();
                case EnumResult.PrintFile.TwoMonth:
                    return (from t in operatorModels
                            where t.Name == MainWindow.poteriashka
                            select t.Proideno31).Sum();
            }
            return (from t in operatorModels
                    where t.Name == MainWindow.poteriashka
                    select t.ProidenoAll).Sum();
        }

        private void PrintTwoPathMonth(string sheetName)
        {
            //День 
            //Пройдено 16-31	
            //Пройдено 1-31	
            //Оклад 
            //Оклад с 16-31	
            //Бонус 
            //Бонусы 16-31	
            //БОНУСЫ с 1-31	
            //Бонусы предыдущего периода 
            //Доплата предыдущего периода 
            //Бонусы к выплате 
            //Премия  
            //Сумма к выплате

            GreateNewSheet(sheetName);
            GreateMotivationSheet();
            var row = 2;
            var col = 1;
#region HeadSheet
            newSheet.SetValue(1, 1, "Оператор");
            newSheet.SetValue(1, 2, "День");
            newSheet.SetValue(1, 3, "Пройдено 16-31");
            newSheet.SetValue(1, 4, "Пройдено 1-31");
            newSheet.SetValue(1, 5, "Оклад");
            newSheet.SetValue(1, 6, "Оклад с 16-31");
            newSheet.SetValue(1, 7, "Бонус");
            newSheet.SetValue(1, 8, "Бонусы 16-31");
            newSheet.SetValue(1, 9, "БОНУСЫ с 1-31");
            newSheet.SetValue(1, 10, "Бонусы предыдущего периода");
            newSheet.SetValue(1, 11, "Доплата предыдущего периода");
            newSheet.SetValue(1, 12, "Бонусы к выплате");
            newSheet.SetValue(1, 13, "Премия");
            newSheet.SetValue(1, 14, "Сумма к выплате");
#endregion
            result = QueryTable(EnumResult.PrintFile.TwoMonth);

            foreach (var cell in result)
            {
                if (cell.Name != MainWindow.poteriashka)
                {
                    newSheet.SetValue(row, col++, cell.Name);
                    newSheet.SetValue(row, col++, cell.Days31);
                    newSheet.SetValue(row, col++, cell.Proideno31);
                    newSheet.SetValue(row, col++, cell.ProidenoAll);
                    newSheet.SetValue(row, col++, cell.Oklad);
                    newSheet.Cells[row, col++].Formula = "=(" + newSheet.Cells[row, 5].Address + "/" + 25 + ")*" + newSheet.Cells[row, 2].Address;
                    newSheet.Cells[row, col++].Formula =
                                                    "IF(" + newSheet.Cells[row, 5].Address + "=" + motivashion.Name + "!" + motivashion.Cells[1, 1].Address
                                                        + ", IF(AND(" + newSheet.Cells[row, 4].Address + ">=" + motivashion.Name + "!" + motivashion.Cells[13, 1].Address + "," + newSheet.Cells[row, 4].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[13, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[13, 3].Address
                                                        + ", IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[14, 1].Address + ", " + newSheet.Cells[row, 4].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[14, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[14, 3].Address
                                                        + ", IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[15, 1].Address + ", " + newSheet.Cells[row, 4].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[15, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[15, 3].Address
                                                        + ", IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[16, 1].Address + ", " + newSheet.Cells[row, 4].Address + "<=" + motivashion.Name + "!" + motivashion.Cells[16, 2].Address + ")," + motivashion.Name + "!" + motivashion.Cells[16, 3].Address
                                                        + ", IF(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[17, 1].Address + "," + motivashion.Name + "!" + motivashion.Cells[17, 3].Address + ",0" + ")))))"
                                                + ", IF(" + newSheet.Cells[row, 5].Address + "=" + motivashion.Name + "!" + motivashion.Cells[1, 5].Address
                                                        + ", IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[4, 5].Address + ", " + newSheet.Cells[row, 4].Address + " <= " + motivashion.Name + "!" + motivashion.Cells[4, 6].Address + "), " + motivashion.Name + "!" + motivashion.Cells[4, 7].Address
                                                        + ", IF(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[5, 5].Address + "," + motivashion.Name + "!" + motivashion.Cells[5, 7].Address + ",0" + ")),0))";
                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 3].Address + "*" + newSheet.Cells[row, 7].Address;
                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 4].Address + "*" + newSheet.Cells[row, 7].Address;
                    newSheet.SetValue(row, col++, cell.BonusDyas15);
                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 9].Address + "-" + newSheet.Cells[row, 10].Address + "-" + newSheet.Cells[row, 8].Address;
                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 8].Address + "+" + newSheet.Cells[row, 11].Address;
                    newSheet.Cells[row, col++].Formula = 
                                                         "IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[4, 9].Address + ", " + newSheet.Cells[row, 4].Address + " <= " + motivashion.Name + "!" + motivashion.Cells[4, 10].Address + "), " + motivashion.Name + "!" + motivashion.Cells[4, 11].Address
                                                     + ", IF(AND(" + newSheet.Cells[row, 4].Address + " >= " + motivashion.Name + "!" + motivashion.Cells[5, 9].Address + ")," + motivashion.Name + "!" + motivashion.Cells[5, 11].Address + ",0))";

                    newSheet.Cells[row, col++].Formula = "=" + newSheet.Cells[row, 6].Address + "+" + newSheet.Cells[row, 12].Address + "+" + newSheet.Cells[row, 13].Address;
                    col = 1;
                    row++;
                }
            }

            newSheet.SetValue(row + 3, 1, "Распределите очки Потеряшки: " + GetPointPoteriashka(EnumResult.PrintFile.TwoMonth));
            StyleTable("A1:N", 14);
        }

        private IEnumerable<OperatorModel> QueryTable(EnumResult.PrintFile type)
        {
            switch (type)
            {
                case EnumResult.PrintFile.FirsMonth:
                    return
                            from t in operatorModels
                            where t.Name != MainWindow.poteriashka && t.Days15 > 0
                            select t;
            }
            return
                    from t in operatorModels
                    where t.Name != MainWindow.poteriashka && t.Days31 > 0
                    select t;
        }

        private void PrintFirstPathMonth(string sheetName)
        {
            GreateNewSheet(sheetName);
            GreateMotivationSheet();
            var row = 2;
            var col = 1;
#region HeadSheet
            newSheet.SetValue(1, 1, "Оператор");
            newSheet.SetValue(1, 2, "День");
            newSheet.SetValue(1, 3, "Пройдено");
            newSheet.SetValue(1, 4, "Оклад");
            newSheet.SetValue(1, 5, "Бонус");
            newSheet.SetValue(1, 6, "Оклад к выплате");
            newSheet.SetValue(1, 7, "Бонусы 1 - 15");
            newSheet.SetValue(1, 8, "Сумма");
#endregion

            result = QueryTable(EnumResult.PrintFile.FirsMonth);
            foreach (var cell in result)
            {
                if (cell.Name != MainWindow.poteriashka)
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
            newSheet.SetValue(row + 3, 1, "Распределите очки Потеряшки: " + GetPointPoteriashka(EnumResult.PrintFile.FirsMonth));
            StyleTable("A1:H", 8);
        }


        private void StyleTable(string startHead, int col)
        {
            for (int i = 1; i <= col; i++)
            {
                newSheet.Column(i).AutoFit(10);
            }

            //var modelRows = operatorModels.Count(x => x.Name != Program.poteriashka)+ 1;

            var modelRows = result.Count() + 1;
            string modelRange = startHead + modelRows.ToString();
            var modelTable = newSheet.Cells[modelRange];

            // Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //Style Head table
            modelTable = newSheet.Cells[startHead + "1"];
            modelTable.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            modelTable.Style.Font.Color.SetColor(Color.DarkGreen);
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }
    }
}