using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;

namespace WhoDidThat
{
    [Transaction(TransactionMode.Manual)]
    public class WhoDidThat : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedElementIds = selection.GetElementIds();

            int counter = 0;
            StringBuilder sb = new StringBuilder();

            foreach (ElementId elementId in selectedElementIds)
            {
                Element element = doc.GetElement(elementId);
                FamilyInstance elementFamily = element as FamilyInstance;
                WorksharingTooltipInfo info = WorksharingUtils.GetWorksharingTooltipInfo(doc, elementId);
                sb.AppendLine("Элемент:    " + element.Name);
                if (elementFamily != null)
                {
                    sb.AppendLine("Семейство: " + elementFamily.Symbol.Family.Name);
                }                
                sb.AppendLine("Создал:    " + info.Creator);
                sb.AppendLine("Изменил:     " + info.LastChangedBy);
                sb.AppendLine("Заемщик:     " + info.Owner);
                sb.AppendLine("  ");

                counter++;
                if (counter % 5 == 0)
                {
                    TaskDialog.Show("Выделение " + (counter - 4).ToString() + "-" + counter.ToString(), sb.ToString());
                    sb.Clear();
                }
            }

            if (counter == 0) sb.Append("Вы ничего не выделили");
            if (sb.Length != 0) TaskDialog.Show("Выделение", sb.ToString());

            return Result.Succeeded;
        }
    }
}
