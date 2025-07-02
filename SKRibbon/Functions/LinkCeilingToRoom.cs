using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Controls;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.Attributes;
namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    class LinkCeilingToRoom : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            // Временно
            Document Doc = doc;

            ICollection<Element> rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();

            ICollection<Element> ceilings = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Ceilings).
                                                    WhereElementIsNotElementType().
                                                    ToElements();

            Transaction t = new Transaction(Doc, "Связать потолки с помещениями");
            t.Start();

            foreach (Element ceiling in ceilings)
            {
                BoundingBoxXYZ boundingBox = ceiling.get_BoundingBox(null);
                XYZ point = new XYZ((boundingBox.Max.X + boundingBox.Min.X) / 2, (boundingBox.Max.Y + boundingBox.Min.Y) / 2, boundingBox.Min.Z - 2);
                foreach (Element room in rooms)
                {
                    Room trueRoom = room as Room;
                    if (trueRoom.IsPointInRoom(point))
                    {
                        Parameter groupParam = ceiling.LookupParameter("ADSK_Группирование");
                        if (groupParam != null) groupParam.Set(trueRoom.Number.ToString());
                        break;
                    }
                }
            }
            t.Commit();

            return Result.Succeeded;
            //throw new NotImplementedException();
        }
    }
}
