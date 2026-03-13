using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using static SKRibbon.FormDesign;
using System.Windows.Controls;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    class CopyParameter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Берем документ
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            //

            ICollection<Element> rooms = new FilteredElementCollector(doc).
                                                    OfCategory(BuiltInCategory.OST_Doors).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
            Transaction t = new Transaction(doc, "Копирование параметра");
            t.Start();
            foreach (Element room in rooms)
            {
                //Parameter nameParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                ElementId elTypeId = room.GetTypeId();
                Element elType = doc.GetElement(elTypeId);

                Parameter nameParam = elType.LookupParameter("Комментарии к типоразмеру");
                Parameter param = elType.LookupParameter("СК_Дверь_Описание");
                if ((param != null) &&(nameParam != null))
                {
                    param.Set(nameParam.AsValueString());
                }
            }
            t.Commit();

            return Result.Succeeded;
        }
    }
}
