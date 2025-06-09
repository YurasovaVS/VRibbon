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
using WinForms = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using Autodesk.Revit.DB.Architecture;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    public partial class NumerateRoomsForm : VForm
    {
        Document Doc;
        FlowLayoutPanel LevelsWrapper = new FlowLayoutPanel();
        FlowLayoutPanel OptionsWrapper = new FlowLayoutPanel();
        WinForms.NumericUpDown LevelNum = new WinForms.NumericUpDown();
        public NumerateRoomsForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;

            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            LevelsWrapper.AutoSize = true;
            LevelsWrapper.FlowDirection = FlowDirection.TopDown;
            LevelsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(LevelsWrapper);
            LevelsWrapper.Anchor = AnchorStyles.Left;

            OptionsWrapper.AutoSize = true;
            OptionsWrapper.FlowDirection = FlowDirection.TopDown;
            OptionsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(OptionsWrapper);
            OptionsWrapper.Anchor = AnchorStyles.Left;

            // Уровни
            // Добавить лейбл
            Label levelsLabel = new Label();
            levelsLabel.Text = "Выберите уровни:";
            levelsLabel.Parent = LevelsWrapper;
            LevelsWrapper.Controls.Add(levelsLabel);

            ICollection<Element> levelsFEC = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Levels).
                WhereElementIsNotElementType().
                ToElements();

            List<Element>levels = levelsFEC.OrderBy(level => level.Name).ToList();

            foreach (Element level in levels)
            {
                LevelCheckBox levelCB = new LevelCheckBox();
                levelCB.level = level;
                levelCB.Text = level.Name;

                levelCB.Parent = LevelsWrapper;
                LevelsWrapper.Controls.Add(levelCB);
                levelCB.Checked = false;

                levelCB.Size = new Size(200, 20);
            }

            // Опции
            Label numLabel = new Label();
            numLabel.Text = "Номер этажа:";
            numLabel.Parent = OptionsWrapper;
            OptionsWrapper.Controls.Add(numLabel);

            LevelNum.Parent = OptionsWrapper;
            OptionsWrapper.Controls.Add(LevelNum);

            LevelNum.Increment = 1;
            LevelNum.Maximum = 100000;
            LevelNum.Minimum = 0;

            Button button = new Button();
            button.Text = "ОК";
            button.Parent = OptionsWrapper;
            OptionsWrapper.Controls.Add(button);
            button.Click += Numerate;

        }

        private void Numerate(object sender, EventArgs e)
        {
            List<ElementXY> roomsToNumerate = new List<ElementXY>();
            for (int i = 1; i<LevelsWrapper.Controls.Count; i++)
            {
                ICollection<Element> rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
                

                LevelCheckBox checkBox = (LevelCheckBox)LevelsWrapper.Controls[i];
                if (checkBox.Checked)
                {
                    foreach (Element room in rooms) {
                        if (room.LevelId == checkBox.level.Id) {                            
                            LocationPoint locationPoint = room.Location as LocationPoint;

                            if (locationPoint == null) continue;

                            double x = locationPoint.Point.X;
                            double y = locationPoint.Point.Y;
                            ElementXY roomXY = new ElementXY(room, x, y);
                            roomsToNumerate.Add(roomXY);
                        }
                    }
                }
            }
            roomsToNumerate = roomsToNumerate.OrderByDescending(room => room.Y). ThenBy(room => room.X).ToList<ElementXY>();

            int a = 1;
            int b = 1;

            foreach (ElementXY roomXY in roomsToNumerate)
            {
                Parameter roomNumParam = roomXY.Room.LookupParameter("Номер");
                if (roomNumParam == null) continue;
                string num = "";
                if (LevelNum.Value == 0)
                {
                    num = a.ToString();
                }
                else {
                    num = a.ToString("00");
                }
                
                if (a <= 99) {
                    num = LevelNum.Value.ToString() + num;
                    a++;
                }
                else
                {
                    num = LevelNum.Value.ToString() + "99/" + b.ToString();
                    b++;
                }
                //bool flag = roomNumParam.SetValueString(num + "a");
                Transaction t = new Transaction(Doc, "Пронумеровать помещения");
                t.Start();
                Room room = (Room)roomXY.Room;
                room.Number = num;
                t.Commit();
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    public class LevelCheckBox : WinForms.CheckBox {
        public Element level;
    }

    public class ElementXY {
        public ElementXY(Element room, double x, double y) { 
            Room = room;
            X = x;
            Y = y;
        }
        public Element Room;
        public double X;
        public double Y;
    }
}
