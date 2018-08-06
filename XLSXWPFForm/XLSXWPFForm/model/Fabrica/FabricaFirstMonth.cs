using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXWPFForm.model.Fabrica
{
    class FabricaFirstMonth: AbstractWorksheet
    {
        ExcelPackage excelIn;
        public override Dictionary<string, AbstractModel> FactoryWorksheet(ExcelPackage excelIn)
        {
            throw new NotImplementedException();
        }
    }
}
