using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm
{
    public class EnumResult
    {
        public enum PrintFile : int
        {
            FirsMonth,
            TwoMonth,
            Motivation,
            AllMonth
        }
        public enum InputOklad : int
        {
            Input,
            Default
        }
    }
}
