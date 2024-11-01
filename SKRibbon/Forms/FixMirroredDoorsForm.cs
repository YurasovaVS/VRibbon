using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;

namespace SKRibbon
{
    public partial class FixMirroredDoorsForm : System.Windows.Forms.Form
    {
        Document Doc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        public FixMirroredDoorsForm(Document doc)
        {
            InitializeComponent();
            this.AutoScroll = true;
            this.Width = 250;
            this.Height = 150;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            Label noDoorsFoundLb = new Label();
            noDoorsFoundLb.Anchor = AnchorStyles.Top;
            noDoorsFoundLb.AutoSize = true;
            noDoorsFoundLb.Parent = formWrapper;
            formWrapper.Controls.Add(noDoorsFoundLb);

            Label mirroredCurtainsFound = new Label();
            mirroredCurtainsFound.Anchor = AnchorStyles.Top;
            mirroredCurtainsFound.AutoSize = true;
            mirroredCurtainsFound.Parent = formWrapper;
            formWrapper.Controls.Add(mirroredCurtainsFound);
            mirroredCurtainsFound.Text = "";

            Button okBtn = new Button();
            okBtn.Anchor = AnchorStyles.Top;
            okBtn.AutoSize = true;
            okBtn.Parent = formWrapper;
            formWrapper.Controls.Add(okBtn);
            okBtn.Click += CloseWindow;
            okBtn.Text = "Спасибо, Витрувий!";

            int doorCount = 0;
            int curtainDoorCount = 0;

            ICollection<Element> doors = new FilteredElementCollector(doc).
                                            OfCategory(BuiltInCategory.OST_Doors).
                                            WhereElementIsNotElementType().
                                            ToElements();
            //List<FamilyInstance> mirroredDoors = new List<FamilyInstance>();
            Transaction t = new Transaction(doc, "Исправление отзеркаленных дверей");
            t.Start();

            foreach (Element door in doors)
            {
                FamilyInstance doorFI = door as FamilyInstance;
                if (doorFI.Mirrored != true)
                {
                    continue;
                }
                //mirroredDoors.Add(doorFI);

                
                /*
                Wall host = doorFI.Host as Wall;

                XYZ normal = host.Orientation;
                XYZ rotatedNormal = new XYZ(-normal.Y, normal.X, normal.Z);

                LocationPoint point = (LocationPoint)door.Location;
                XYZ origin = new XYZ(point.Point.X, point.Point.Y, point.Point.Z);

                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(rotatedNormal, origin);

                ElementTransformUtils.MirrorElement(doc, door.Id, mirrorPlane);
                */

                LocationPoint point = (LocationPoint)door.Location;
                if (point == null) {
                    curtainDoorCount++;
                    continue;
                }
                doorCount++;
                Level level = doc.GetElement(door.LevelId) as Level;
                XYZ xyz = new XYZ(point.Point.X, point.Point.Y, point.Point.Z);
                FamilySymbol symbol = doorFI.Symbol;
                Element host = doorFI.Host;
                StructuralType structuralType = doorFI.StructuralType;
                bool flag = doorFI.FacingFlipped;

                doc.Delete(door.Id);
                FamilyInstance mirroredDoor = doc.Create.NewFamilyInstance(xyz, symbol, host, level, structuralType);
                if (flag)
                {
                    mirroredDoor.rotate();
                }
            }

            t.Commit();
                      

            if (doorCount <= 0)
            {

                noDoorsFoundLb.Text = "Отзеркаленные двери не найдены.";
            }
            else
            {
                noDoorsFoundLb.Text = "Найдено и исправлено " + doorCount.ToString() + " дверей.";

            }

            if (curtainDoorCount > 0) 
            {
                mirroredCurtainsFound.Text = "Обнаружено " + curtainDoorCount.ToString() + " отзеркаленных дверей в витражах. Эти двери не исправлены!";
                this.Width = 480;
                this.Height = 150;
            }

        }

        public void CloseWindow(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
