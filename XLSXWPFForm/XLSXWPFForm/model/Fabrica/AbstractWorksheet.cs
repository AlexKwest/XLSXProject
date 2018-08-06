using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model.Fabrica
{
    public abstract class AbstractWorksheet
    {
        ExcelPackage excelIn;
        public abstract Dictionary<string, AbstractModel> FactoryWorksheet(ExcelPackage excelIn);
    }
}
