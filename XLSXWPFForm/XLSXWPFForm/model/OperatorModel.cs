using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model
{
    public class OperatorModel
    {
        public override string ToString()
        {
            return Name;
        }
        public string ResultSummOperator
        {
            get
            {
                return $"{Summa} руб./{Summa31} руб.";
            }
        }

        private int premia;
        public string Name { get; set; }
        public int Days15 { get; set; } 
        public int Proideno15 { get; set; }  
        public int Oklad { get; set; } 
        public int Bonus15
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (Proideno15 >=1 && Proideno15 <= 4)
                        return 200;
                    if (Proideno15 >= 5 && Proideno15 <= 9)
                        return 400;
                    if (Proideno15 >= 10 && Proideno15 <= 14)
                        return 600;
                    if (Proideno15 >= 15)
                        return 1000;
                }
                if (Oklad > 6250)
                {
                    if (Proideno15 >=1 && Proideno15 <= 10)
                        return 50;
                    if (Proideno15 >=11)
                        return 100;
                }
                return 0;
            }
        }
        public float OkladinPay15
        {
            get
            {
                return Oklad / 25 * Days15;
            }
        }
        public int BonusDyas15
        {
            get
            {
                return Proideno15 * Bonus15;
            }
        }
        public float Summa
        {
            get
            {
                return OkladinPay15 + BonusDyas15;
            }
        }
        public string Show()
        {
            return ($"Имя - {Name}: Дней - {Days15}; Пройдено - {Proideno15}: Оклад - {Oklad};  Бонус - {Bonus15};" +
                    $" Оклад к оплате - {OkladinPay15}; Бонус за 15 - {BonusDyas15};" +
                    $" ----> {Summa} " + "\n" +
                    "----------------------------------------------------------------------------------------------------------------------------------");
        }

        public int Days31 { get; set; } 
        public int Proideno31 { get; set; }
        public int ProidenoAll
        {
            get
            {
                return Proideno15 + Proideno31;
            }
        }
        public float OkladinPay31
        {
            get
            {
                return Oklad / 25 * Days31;
            }
        }
        public int Bonus31
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (ProidenoAll >=1 && ProidenoAll <= 4)
                        return 200;
                    if (ProidenoAll >= 5 && ProidenoAll <= 9)
                        return 400;
                    if (ProidenoAll >= 10 && ProidenoAll <= 14)
                        return 600;
                    if (ProidenoAll >= 15 && ProidenoAll <= 19)
                        return 800;
                    if (ProidenoAll >= 20)
                        return 1000;
                }
                else if (Oklad > 6250)
                {
                    if (ProidenoAll >=1 && ProidenoAll <= 10)
                        return 50;
                    if (ProidenoAll >= 11)
                        return 100;
                }
                return 0;
            }
        }
        public int BonusDyas31
        {
            get
            {
                return Proideno31 * Bonus31;
            }
        }                                                                
        public int BonusDyasAll
        {
            get
            {
                return ProidenoAll * Bonus31;
            }
        } 
        public int DoplataLastPeriod
        {
            get
            {
                return BonusDyasAll - BonusDyas15 - BonusDyas31;
            }
        }       
        public int BonusInPay
        {
            get
            {
                return DoplataLastPeriod + BonusDyas31;
            }
        }           
        public int Premia 
        {
            get
            {
                if (ProidenoAll >= 30 && ProidenoAll <=34)
                {
                    return 3000;
                }
                else if (ProidenoAll >= 35)
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
        public float Summa31
        {
            get
            {
                return BonusInPay + Premia + OkladinPay31;
            }
        }
    }
}
