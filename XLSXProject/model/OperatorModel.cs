using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXProject.model
{
    public class OperatorModel : IEqualityComparer<OperatorModel>
    {
       // private int proideno15;
       // private int proideno31;
        private int premia;

        //public static int Poteriashki = 0 ;
        public string Name { get; set; }
        public int Days15 { get; set; } // BJ18
        public int Proideno15 { get; set; }  //BL18
        //{
        //    get
        //    {
        //        return proideno15;
        //    }
        //    set
        //    {
        //        if ( Name == "потеряшки")
        //        {
        //           // proideno15 = Convert.ToInt32(value);
        //            OperatorModel.Poteriashki += value;
        //        }
        //        proideno15 = value;
        //    }
        //}
        public int Oklad { get; set; } //10 000 6250
        public int Bonus15
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

        //  Name                                                                                                    Имя
        public int Days31 { get; set; } //DZ3   //=BN3+BR3+BV3+BZ3+CD3+CH3+CL3+CP3+CT3+CX3+DB3+DF3+DJ3+DN3+DR3+DV3  День
        public int Proideno31 { get; set; }
        //{
        //    get
        //    {
        //        return proideno31;
        //    }
        //    set
        //    {
        //        if (Name == "потеряшки")
        //        {
        //            // proideno15 = Convert.ToInt32(value);
        //            OperatorModel.Poteriashki += value;
        //        }
        //        proideno31 = value;
        //    }
        //}      //EB3   //=BP3+BT3+BX3+CB3+CF3+CJ3+CN3+CR3+CZ3+DD3+DH3+DL3+DP3+DT3+DX3+CV3  Проидено 16-31
        public int ProidenoAll
        {
            get
            {
                return Proideno15 + Proideno31;
            }
        }       //EF3   //=BL3+EB3                                                          Пройдено 1-31
        //  Oklad                                                                                                   Оклад
        public float OkladinPay31       //=(E2/25)*B2                                                               Оклад с 16-31
        {
            get
            {
                return Oklad / 25 * Days31;
            }
        }    
        public int Bonus31 //                                                                                       Бонус
        {
            get
            {
                if (Oklad <= 6250)
                {
                    if (Proideno31 <= 4)
                        return 200;
                    if (Proideno31 <= 9)
                        return 400;
                    if (Proideno31 <= 14)
                        return 600;
                    else return 1000;
                }
                if (Oklad > 6250)
                {
                    if (Proideno31 <= 10)
                        return 50;
                    else return 100;
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
        }       //=G2*C2                                                                    Бонусы 16-31
        public int BonusDyasAll
        {
            get
            {
                return ProidenoAll * Bonus31;
            }
        } //                                                                               БОНУСЫ с 1-31
        //  public int BonusDyas15                                                                                  Бонусы предыдущего периода
        public int DoplataLastPeriod
        {
            get
            {
                return BonusDyasAll - BonusDyas15 - BonusDyas31;
            }
        }       //=I2-J2-H2                                                           Доплата предыдущего периода
        public int BonusInPay
        {
            get
            {
                return DoplataLastPeriod + BonusDyas31;
            }
        }             //=H2+K2                                                               Бонусы к выплате
        public int Premia //TODO :                                                                                  Премия
                          //=ЕСЛИ(И(D2>=Мотивация!$J$6;D2<=Мотивация!$K$6);Мотивация!$L$6;ЕСЛИ(И(D2>=Мотивация!$J$7);Мотивация!$L$7;"0"))
        {
            get
            {
                return premia;
            }
            set
            {
                premia = value;
            }
        }
        public float Summa31 //=L2+M2+F2                                                                            Сумма к выплате
        {
            get
            {
                return BonusInPay + Premia + OkladinPay31;
            }
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
