using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model
{
    public abstract class AbstractModel
    {
        public override string ToString()
        {
            return Name;
        }
        //Оператор
        public virtual string Name { get; set; }
        //День
        public virtual int Day { get; set; }
        //Пройдено
        public virtual int Finished { get; set; }
        //Оклад
        public virtual int Oklad { get; set; }
        //Bonus
        public abstract int Bonus { get; set; }
        //Бонусы с х до у
        private int bonusDyas;
        public virtual int BonusDyas
        {
            get
            {
                return Finished * Bonus;
            }

            set
            {
                bonusDyas = value;
            }
        }
        //Оклад к выплате
        private int okladInPay;
        public virtual int OkladInPay
        {
            get
            {
                return Oklad / 25 * Day;
            }

            set
            {
                okladInPay = value;
            }
        }
        //Сумма
        public abstract int Summa { get; set; }
    }
}
