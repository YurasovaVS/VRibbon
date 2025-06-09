/*
 * --------------------------------------------------------------------------------------
 * "Витрувий" (Vitruvius) - бесплатный плагин для Autodesk(c) Revit(c), 
 * предназначенный для автоматизации рутинных задач и упрощения работы архитекторов.
 * 
 * Copyright (C) 2023-2025 Юрасова В.С. 
 * 
 * Данная программа относится к категории свободного программного обеспечения.
 * Вы можете распространять и/или модифицировать её согласно условиям Стандартной
 * Общественной Лицензии GNU, опубликованной Фондом Свободного Программного
 * Обеспечения, версии 3.
 * http://www.gnu.org/licenses/.
 * 
 * -------------------------------------------------------------------------------------- * 
 * "Vitruvius" is a free plugin for Autodesk(c) Revit(c), aimed to automate
 * routine tasks and make life easier for architects.
 * 
 * Copyright (C) 2023-2025 Yurasova V.S.
 * 
 *  This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License. * 
 * 
 *  <https://www.gnu.org/licenses/>.
 * 
 * --------------------------------------------------------------------------------------
 */

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
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    public partial class FixMirroredDoorsForm : VForm
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
