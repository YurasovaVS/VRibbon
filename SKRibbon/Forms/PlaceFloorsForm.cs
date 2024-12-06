using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExpIFCUtils = Autodesk.Revit.DB.IFC.ExporterIFCUtils;
using System.Windows.Forms;
using System.Windows.Controls;

namespace SKRibbon
{
    public partial class PlaceFloorsForm : WinForms.Form
    {
        Document Doc;
        List<Element> FloorTypes;
        WinForms.ComboBox floorTypesCB = new WinForms.ComboBox();
        public PlaceFloorsForm(Document doc, ICollection<ElementId> selectionIds)
        {
            InitializeComponent();
            Doc = doc;

            // Обертка для всей формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Обертка для главных настроек (Левая сторона)

            FlowLayoutPanel optionsWrapper = new FlowLayoutPanel();
            optionsWrapper.AutoSize = true;
            optionsWrapper.FlowDirection = FlowDirection.TopDown;
            optionsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(optionsWrapper);
            optionsWrapper.Anchor = AnchorStyles.Left;

            // Выбор комнат (Selection, Active View, Entire project)
            WinForms.Label selectionLabel = new WinForms.Label();
            selectionLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(selectionLabel);
            selectionLabel.Text = "Выбор помещений:";
            selectionLabel.Size = new Size(300, 20);

            WinForms.ComboBox selectionCB = new WinForms.ComboBox();
            selectionCB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(selectionCB);
            selectionCB.Items.Add("Во всем проекте");
            selectionCB.Items.Add("На текущем виде");
            if (selectionIds.Count > 0) {
                selectionCB.Items.Add("Выбранные");
            }
            selectionCB.SelectedIndex = 0;
            selectionCB.Size = new Size(300, 20);


            // Оффсет от уровня (Только цифры)

            WinForms.Label offsetLabel = new WinForms.Label();
            offsetLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(offsetLabel);
            offsetLabel.Text = "Уровень пола от уровня этажа:";
            offsetLabel.Size = new Size(300, 20);
            offsetLabel.Padding = new Padding(0, 10, 0, 0);

            WinForms.TextBox offsetTB = new WinForms.TextBox();
            offsetTB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(offsetTB);
            offsetTB.Text = "0";
            offsetTB.Size = new Size(300, 20);
            offsetTB.Padding = new Padding(0, 5, 0, 0);

            // Пол по умолчанию
            WinForms.Label floorTypesLabel = new WinForms.Label();
            floorTypesLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(floorTypesLabel);
            floorTypesLabel.Text = "Пол по умолчанию:";
            floorTypesLabel.Size = new Size(300, 20);
            floorTypesLabel.Padding = new Padding(0, 10, 0, 0);

            
            floorTypesCB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(floorTypesCB);
            floorTypesCB.Size = new Size(300, 30);
            floorTypesCB.Padding = new Padding(0, 5, 0, 0);

            ICollection<Element> floorTypes = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Floors).
                                                    WhereElementIsElementType().
                                                    ToElements();
            FloorTypes = floorTypes.ToList();

            foreach (Element floorType in FloorTypes) {
                floorTypesCB.Items.Add(floorType.Name);
            }

            floorTypesCB.SelectedIndex = 0;

            //Кнопка

            WinForms.Button button = new WinForms.Button();
            button.Text = "OK";

            button.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(button);
            button.Click += CreateFloors;

            // Обертка для настроек пола (Правая сторона)

            // Будет добавлено позже


            FloorTypes = floorTypes.ToList();
        }

        private void CreateFloors(object sender, EventArgs e)
        {


            ICollection<Element> rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();

            Transaction t = new Transaction(Doc, "Создать полы");
            t.Start();

            //Element room = rooms.FirstOrDefault<Element>();
            foreach (Element room in rooms)
            {
                SpatialElement roomSE = room as SpatialElement;
                SpatialElementBoundaryOptions boundaryOptions = new SpatialElementBoundaryOptions();
                CurveArray curveArray = new CurveArray();

                IList<CurveLoop> curveLoops = ExpIFCUtils.GetRoomBoundaryAsCurveLoopArray(roomSE, boundaryOptions, true);
           
                foreach (Curve curve in curveLoops[0]) {
                    curveArray.Append(curve);
                }

                int index = floorTypesCB.SelectedIndex;


                FloorType floorType = FloorTypes[index] as FloorType;
                Level level = Doc.GetElement(room.LevelId) as Level;
                Doc.Create.NewFloor(curveArray, floorType, level, false, XYZ.BasisZ);
                
                //SketchPlane sp = SketchPlane.Create(Doc, level.Id);
                //foreach (Curve c in curveArray) Doc.Create.NewModelCurve(c, sp);
            }

            t.Commit();

            this.DialogResult = DialogResult.OK;
            this.Close();
                
        }
    }
}
