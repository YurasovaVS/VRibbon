using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using FakeArea;

namespace AddSignatures
{
    [Transaction(TransactionMode.Manual)]
    public class AddSignaturesDWG : IExternalCommand
    {
        public struct Signature 
        {
            public string path;
            public ImportInstance sigInstance;

            public Signature(string path, ImportInstance instance)
            {
                this.path = path;
                this.sigInstance = instance;
            }
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Берем документ
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            //Прописываем путь к подписям
            string path = "\\\\ABSKNAS\\переезд\\13_Пользователи\\_ПодписиСК\\";

            using (System.Windows.Forms.Form form = new SKRibbon.AddSigForm(doc, path))
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }
            //ДОПОЛНИТЕЛЬНО - реализовать выбор пути к подписям

            /*
            DWGImportOptions importOptions = new DWGImportOptions();

            //ДОПОЛНИТЕЛЬНО - реализовать выбор листов, на которых необходимо проставить подписи

            
            ICollection<Element> sheets = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_Sheets)
                                    .WhereElementIsNotElementType()
                                    .ToElements();
         
            
            StringBuilder sb = new StringBuilder();
            
            //Открываем транзакцию
            Transaction t = new Transaction(doc, "Вставить подписи");
            t.Start();
            //Для каждого листа из списка листов:
            foreach (ViewSheet sheet in sheets)
            {
                // Определяем координаты правого нижнего угла листа
                Double sheetRBpointX = sheet.Outline.Max.U;
                Double sheetRBpointY = sheet.Outline.Min.V;

                for (int i = 0; i <= 5; i++)
                {
                    string paramName = "ADSK_Штамп Строка " + (i + 1).ToString() + " фамилия";
                    string paramValue = sheet.LookupParameter(paramName).AsString();

                    // Если фамилия задана
                    if (paramValue != null)
                    {
                        if (paramValue.Length != 0)
                        {
                            string signaturePath = path + "подпись_" + paramValue + ".dwg";
                            if (File.Exists(signaturePath))
                            {
                                ElementId signatureId;
                                bool flag = doc.Link(signaturePath, importOptions, sheet, out signatureId);
                                Element signature = doc.GetElement(signatureId);
                                signature.Pinned = false;
                                Double x = sheetRBpointX - 0.49;
                                Double y = 30 - i * 5;
                                y = UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters);
                                y = sheetRBpointY + y;
                                XYZ move = new XYZ(x, y, 0);
                                signature.Location.Move(move);
                            }
                            else
                            {
                                sb.AppendLine("Подпись не найдена: " + signaturePath);
                            }
                        }
                    }
                }

            }
            if (sb.Length == 0)
            {
                sb.AppendLine("Подписи проставлены. Ошибок нет.");
            }
            TaskDialog.Show("Ошибки", sb.ToString());
            
            t.Commit();
            */
            return Result.Succeeded;
        }
    }
}
