using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;

namespace FilterByPeople
{
    [Transaction(TransactionMode.Manual)]
    public class FilterByPeople : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiApp.ActiveUIDocument.Document;

            ICollection<ElementId> elId = uiDoc.Selection.GetElementIds();
            if (uiDoc.Selection.GetElementIds().Count != 0)
            {
                using (System.Windows.Forms.Form form = new FilterByPeopleForm(uiDoc))
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
            } else
            {
                TaskDialog.Show("Ошибка", "Вы ничего не выделили");
                return Result.Failed;
            }
            
        }
    }
}
