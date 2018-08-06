using System.Collections.Generic;
using System.Linq;
using XLSXWPFForm.model;

namespace XLSXWPFForm
{
    public static class UnionDictynary
    {
        static List<OperatorFirstMonthModel> firstExcel_List;
        static List<OperatorTwoMonthModel> twoExcel_List;

        public static void UnionResult (List<OperatorFirstMonthModel> firstExcel_List, List<OperatorTwoMonthModel> twoExcel_List)
        {
            UnionDictynary.firstExcel_List = firstExcel_List;
            UnionDictynary.twoExcel_List = twoExcel_List;
        }

        public static Dictionary<string, OperatorTwoMonthModel> Result( )
        {
            var resultFirst = firstExcel_List.ToDictionary(x => x.Name, x => x);
            var twoExcel = twoExcel_List.ToDictionary(x => x.Name, x => x);
            Dictionary<string, OperatorTwoMonthModel> result = new Dictionary<string, OperatorTwoMonthModel>();
            //Cравниваем и сращиваем два списка. Базовый и первого месяца по имени.
            foreach (var bases in twoExcel_List)
            {
                var flag = false;
                foreach (var first in firstExcel_List)
                {
                    if (bases.Name == first.Name)
                    {
                        flag = true;
                        result.Add(bases.Name,
                            new OperatorTwoMonthModel
                            {
                                Name = bases.Name,
                                Day = bases.Day,
                                Finished = bases.Finished,
                                FinishedLast = first.Finished,
                                // Пройдено 1-31 FinishedAll
                                Oklad = first.Oklad,
                                // Оклад к выплате OkladInPay
                                // Бонус Bonus
                                // БОНУСЫ с 1-31 BonusAll
                                Bonus15 = first.BonusDyas,
                                //Доплата предыдущего периода DoplataLast
                                // Бонусы к выплате BonusInPay
                                // Премия Premia
                                //Summa
                            }
                            );
                    }
                }
                if (!flag)
                {
                    result.Add(bases.Name, bases);
                }
            }
            return result;
            #region Commit
            //first.Add("1", new List<string> { "11", "12" });
            //first.Add("2", new List<string> { "1", "2" });
            //second.Add("2", new List<string> { "2", "3","1","4" });
            //second.Add("3", new List<string> { "21", "31" });
            //var d1 = first.Union(second).GroupBy(pair => pair.Key).
            //ToDictionary(gr => gr.Key, gr => gr.SelectMany(pair => pair.Value).Distinct().ToList());
            ////другой вариант
            //first1.Where(pair => !second1.ContainsKey(pair.Key)).Union(second1).All(pair => { RES1.Add(pair.Key, pair.Value); return true; });

            //result = twoExcel.Where(pair => !resultFirst.ContainsKey(pair.Key)) //.Union(twoExcel).All(pair => { result.Add(pair.Key, pair.Value); return true; });
            #endregion
        }
    }
}
