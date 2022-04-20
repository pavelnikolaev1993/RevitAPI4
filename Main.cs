using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    string pipeName = pipe.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString();
                    string pipeOuter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsString();
                    string pipeInner = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsString();
                    string pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsString();
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeName);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipeOuter);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipeInner);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipeLength);
                    rowIndex++;
                }
                workbook.Write(stream);
                workbook.Close();
            }
            System.Diagnostics.Process.Start(excelPath);
            return Result.Succeeded;
        }
    }
}
