using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model
{
    //Оператор	День	Пройдено 16-31	Пройдено 1-31	Оклад	Оклад с 16-31	Бонус	Бонусы 16-31	БОНУСЫ с 1-31	Бонусы предыдущего периода	Доплата предыдущего периода	Бонусы к выплате	Премия	Сумма к выплате

    public class OperatorTwoMonthModel : AbstractModel
    {
        //Оператор Name
        // День Day
        //Пройдено с 1-16
        public int FinishedLast { get; set; }
        //Пройдено 1-31
        public int FinishedAll
        {
            get
            {
                return FinishedLast + Finished;
            }
        }
        //Оклад
        //Оклад с 16-31
        //Бонус
        private int bonus;
        public override int Bonus
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (FinishedAll >= 1 && FinishedAll <= 4)
                        return 200;
                    if (FinishedAll >= 5 && FinishedAll <= 9)
                        return 400;
                    if (FinishedAll >= 10 && FinishedAll <= 14)
                        return 600;
                    if (FinishedAll >= 15 && FinishedAll <= 19)
                        return 800;
                    if (FinishedAll >= 20)
                        return 1000;
                }
                else if (Oklad > 6250)
                {
                    if (FinishedAll >= 1 && FinishedAll <= 10)
                        return 50;
                    if (FinishedAll >= 11)
                        return 100;
                }
                return 0;
            }
            set
            {
                bonus = value;
            }
        }
        //БОНУСЫ с 1-31
        public int BonusAll
        {
            get
            {
                return FinishedAll * Bonus;
            }
        }
        //Бонусы предыдущего периода
        public int Bonus15 { get; set; }
        //Доплата предыдущего периода
        public int DoplataLast
        {
            get
            {
                return BonusAll - Bonus15 - BonusDyas;
            }
        }
        //Бонусы к выплате
        public int BonusInPay
        {
            get
            {
                return BonusDyas + DoplataLast;
            }
        }
        //Премия
        private int premia;
        public int Premia
        {
            get
            {
                if (FinishedAll >= 30 && FinishedAll <= 34)
                {
                    return 3000;
                }
                else if (FinishedAll >= 35)
                {
                    return 5000;
                }
                else return 0;
            }
            set
            {
                premia = value;
            }
        }
        //Сумма к выплате
        private int summa;
        public override int Summa
        {
            get
            {
                return BonusInPay + Premia + OkladInPay;
            }
            set
            {
                summa = value;
            }
        }

        public double PercentPremia
        {
            get
            {
                var result = (double)100 / 35 * FinishedAll;
                return result > 100 ? 100 : result;
            }
        }

    }
}
