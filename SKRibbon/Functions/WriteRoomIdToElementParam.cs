using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SheetRenamer;
using SKRibbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    internal class WriteRoomIdToElementParam : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            // Выделение для передачи в форму
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Selection selection = uidoc.Selection;
            ICollection<ElementId> selectionIds = uidoc.Selection.GetElementIds();

            using (SKRibbon.FormDesign.VForm form = new WriteRoomIdToElementParamForm(doc, selectionIds))
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
        }
    }
}
