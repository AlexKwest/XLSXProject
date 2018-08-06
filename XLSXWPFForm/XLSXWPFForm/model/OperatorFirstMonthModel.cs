using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model
{
    public class OperatorFirstMonthModel: AbstractModel
    {
        //Оператор
        //День
        //Пройдено
        //Оклад
        //Бонус
        private int bonus;
        public override int Bonus
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (Finished >= 1 && Finished <= 4)
                        return 200;
                    if (Finished >= 5 && Finished <= 9)
                        return 400;
                    if (Finished >= 10 && Finished <= 14)
                        return 600;
                    if (Finished >= 15)
                        return 1000;
                }
                if (Oklad > 6250)
                {
                    if (Finished >= 1 && Finished <= 10)
                        return 50;
                    if (Finished >= 11)
                        return 100;
                }
                return 0;
            }
            set
            {
                bonus = value;
            }
        }
        //Оклад к выплате
        //Бонусы 1 - 15
        //Сумма
        private int summa;
        public override int Summa
        {
            get
            {
               return OkladInPay + BonusDyas;
            }
            set
            {
                summa = value;
            }
        }
    }
}
