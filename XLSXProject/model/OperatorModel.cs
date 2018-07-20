using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXProject.model
{
    public class OperatorModel : IEqualityComparer<OperatorModel>
    {
        private int days15;
        private int proideno15;
        public static int Poteriashki = 0 ;
        //private string name;

        public string Name { get; set; }

        public int Days15 // BJ18
        {
            get
            {
                return days15;
            }
            set
            {
                days15 = Convert.ToInt32(value);
            }
        }

        public int Proideno15  //BL18
        {
            get
            {
                return proideno15;
            }
            set
            {
                if ( Name == "потеряшки")
                {
                   // proideno15 = Convert.ToInt32(value);
                    OperatorModel.Poteriashki += value;
                }
                proideno15 = value;
            }
        }

        public int Oklad { get; set; } //10 000 6250
        public int Bonus
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (Proideno15 <= 4)
                        return 200;
                    if (Proideno15 <= 9)
                        return 400;
                    if (Proideno15 <= 14)
                        return 600;
                    else return 1000;
                }
                if (Oklad > 6250)
                {
                    if (Proideno15 <= 10)
                        return 50;
                    else return 100;
                }
                return 0;
            }
        }
        public float OkladinPay
        {
            get
            {
                return Oklad / 25 * Proideno15;
            }
        }
        public int BonusDyas15
        {
            get
            {
                return Proideno15 * Bonus;
            }
        }
        public float Summa
        {
            get
            {
                return OkladinPay + BonusDyas15;
            }
        }

        public string Show()
        {
            return ($"Имя - {Name}: Дней - {Days15}; Пройдено - {Proideno15}: Оклад - {Oklad};  Бонус - {Bonus};" + 
                    $" Оклад к оплате - {OkladinPay}; Бонус за 15 - {BonusDyas15};" + 
                    $" ----> {Summa} " + "\n" +
                    "----------------------------------------------------------------------------------------------------------------------------------");
        }


        public bool Equals(OperatorModel x, OperatorModel y)
        {
            return x.Name == y.Name;
        }

        public int GetHashCode(OperatorModel obj)
        {
            return obj.Name.GetHashCode();
        }



    }
}
