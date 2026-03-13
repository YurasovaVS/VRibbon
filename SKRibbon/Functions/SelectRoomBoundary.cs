using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    class SelectRoomBoundary : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Берем документ
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;

            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedElementIds = selection.GetElementIds();
            ICollection<Element> rooms = new FilteredElementCollector(doc, selectedElementIds).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
            

            HashSet<ElementId> boundaryElementsIds = new HashSet<ElementId>();

            foreach (Element roomEl in rooms)
            {
                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

                Room room = roomEl as Room;
                //GetGeneratingElementIds
                foreach (IList<Autodesk.Revit.DB.BoundarySegment> boundSegList in room.GetBoundarySegments(options))
                {
                    foreach (Autodesk.Revit.DB.BoundarySegment boundSeg in boundSegList)
                    {
                        boundaryElementsIds.Add(boundSeg.ElementId);                        
                    }
                }

                

            }
            uiDoc.Selection.SetElementIds(boundaryElementsIds);
            Transaction t = new Transaction(doc, "Изолировать выделение");
            t.Start();
            doc.ActiveView.IsolateElementsTemporary(uiDoc.Selection.GetElementIds());
            t.Commit();

            return Result.Succeeded;
        }
    }
}
