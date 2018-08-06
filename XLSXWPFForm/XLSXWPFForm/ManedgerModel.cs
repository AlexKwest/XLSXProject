using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXWPFForm.model;
using XLSXWPFForm.model.Fabrica;
namespace XLSXWPFForm
{
    public class ManedgerModel
    {
        ExcelPackage excelIn;
        ExcelPackage excelOut;
        ExcelWorksheet worksheet;

        ExcelWorksheet motivashion;
        ExcelWorksheet newSheet;

        private AbstractWorksheet abstractWorksheet;
        private Dictionary<string, AbstractModel> operatorModels;

        public ManedgerModel (AbstractWorksheet abstractWorksheet , ExcelPackage excelIn)
        {
            this.excelIn = excelIn;
            this.abstractWorksheet = abstractWorksheet;

            excelOut = new ExcelPackage();
            operatorModels = abstractWorksheet.FactoryWorksheet(excelIn);
        }

        public Dictionary<string, AbstractModel> Run()
        {
            return operatorModels;
        }
    }
}
