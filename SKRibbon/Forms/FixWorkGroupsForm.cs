using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRibbon
{
    
    public partial class FixWorkGroupsForm : System.Windows.Forms.Form
    {
        Dictionary<BuiltInCategory, string> CategoryNames = new Dictionary<BuiltInCategory, string>()
        {
            { BuiltInCategory.OST_CurtaSystem, "Витражи"},
            { BuiltInCategory.OST_Doors, "Двери"},
            { BuiltInCategory.OST_Roofs, "Кровля"},
            { BuiltInCategory.OST_Stairs, "Лестницы"},
            //{ BuiltInCategory.OST_, "Оборудование"},
            { BuiltInCategory.OST_Railings, "Ограждения"},
            { BuiltInCategory.OST_Windows, "Окна"},
            //{ BuiltInCategory.OST_, "Отверстия"},
            { BuiltInCategory.OST_Floors, "Полы"},
            { BuiltInCategory.OST_Rooms, "Помещения"},
            { BuiltInCategory.OST_Ceilings, "Потолки"}
        };
        Document Doc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        public FixWorkGroupsForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            IEnumerable<Workset> worksets = new FilteredWorksetCollector(Doc).OfKind(WorksetKind.UserWorkset).ToWorksets();
            ICollection<Element> elements = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements();
        }
    }

}
