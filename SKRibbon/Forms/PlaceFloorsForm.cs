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
using ExpIFCUtils = Autodesk.Revit.DB.IFC.ExporterIFCUtils;
using System.Windows.Forms;
using System.Windows.Controls;
using Autodesk.Revit.DB.Architecture;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    public partial class PlaceFloorsForm : VForm
    {
        Document Doc;
        List<Element> FloorTypes;
        ICollection<Element> floorTypes;
        WinForms.ComboBox floorTypesCB = new WinForms.ComboBox();
        WinForms.ComboBox selectionCB = new WinForms.ComboBox();
        WinForms.NumericUpDown offsetTB = new WinForms.NumericUpDown();
        FlowLayoutPanel settingsWrapper = new FlowLayoutPanel();
        ICollection<ElementId> SelectionIds;
        List<string> roomTypes = new List<string>();

        FunctionMode Mode;


        public PlaceFloorsForm(Document doc, ICollection<ElementId> selectionIds, string mode)
        {
            InitializeComponent();
            Doc = doc;
            SelectionIds = selectionIds;


            switch (mode) {
                case "Полы":
                    Mode = new FunctionMode(BuiltInCategory.OST_Floors, "полы");
                    break;
                case "Потолки":
                    Mode = new FunctionMode(BuiltInCategory.OST_Ceilings, "потолки");
                    break;

            }

            // Обертка для всей формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Опции (панель слева)
            // Обертка для опций
            FlowLayoutPanel optionsWrapper = new FlowLayoutPanel();
            optionsWrapper.AutoSize = true;
            optionsWrapper.FlowDirection = FlowDirection.TopDown;
            optionsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(optionsWrapper);
            optionsWrapper.Anchor = AnchorStyles.Top;
            optionsWrapper.Margin = new Padding(5, 0, 0, 0);
            optionsWrapper.BorderStyle = BorderStyle.FixedSingle;

            int leftPanelWidth = 300;
            // Заголовок
            /*
            VHeaderLabel optionsHeader = new VHeaderLabel();
            optionsHeader.Text = "НАСТРОЙКИ";
            optionsHeader.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(optionsHeader);
            */
            // Опция 1. Выбор помещения (Selection, Active View, Entire project)
            // Заголовок
            WinForms.Label selectionLabel = new WinForms.Label();
            selectionLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(selectionLabel);
            selectionLabel.Text = "Выбор помещений:";
            selectionLabel.Size = new Size(leftPanelWidth, 30);
            selectionLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);

            // Выпадающий список
            selectionCB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(selectionCB);
            selectionCB.Items.Add("Во всем проекте");
            selectionCB.Items.Add("На текущем виде");
            selectionCB.SelectedIndex = 0;
            if (selectionIds.Count > 0) {
                selectionCB.Items.Add("Выбранные");
                selectionCB.SelectedIndex = 2;
            }
            selectionCB.Size = new Size(leftPanelWidth, 30);

            // Опция 2. Оффсет от уровня этажа
            // Заголовок
            WinForms.Label offsetLabel = new WinForms.Label();
            offsetLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(offsetLabel);
            offsetLabel.Text = "Отступ от уровня этажа:";
            offsetLabel.Size = new Size(leftPanelWidth, 30);
            offsetLabel.Padding = new Padding(0, 10, 0, 0);
            offsetLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);

            // Цифры
            offsetTB.Increment = 1;
            offsetTB.Maximum = 100000;
            offsetTB.Minimum = -100000;
            
            offsetTB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(offsetTB);
            offsetTB.Text = "0";
            offsetTB.Size = new Size(leftPanelWidth, 30);
            offsetTB.Padding = new Padding(0, 5, 0, 0);

            // Опция 3. Пол по умолчанию
            // Заголовок
            WinForms.Label floorTypesLabel = new WinForms.Label();
            floorTypesLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(floorTypesLabel);
            floorTypesLabel.Text = "Тип по умолчанию:";
            floorTypesLabel.Size = new Size(leftPanelWidth, 30);
            floorTypesLabel.Padding = new Padding(0, 10, 0, 0);
            floorTypesLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);

            // Выпадающий список
            floorTypesCB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(floorTypesCB);
            floorTypesCB.Size = new Size(leftPanelWidth, 30);
            floorTypesCB.Padding = new Padding(0, 5, 0, 0);

            floorTypes = new FilteredElementCollector(Doc).
                                                    OfCategory(Mode.BuiltInCategory).
                                                    WhereElementIsElementType().
                                                    ToElements();
            FloorTypes = floorTypes.OrderBy(p=>p.Name).ToList();

            foreach (Element floorType in FloorTypes) {
                floorTypesCB.Items.Add(floorType.Name);
            }

            floorTypesCB.SelectedIndex = 0;

            //Кнопка

            VButton button = new VButton();
            button.Text = "ЗАМЕНИТЬ ПОЛЫ";
            button.Size = new Size(leftPanelWidth, 50);
            //button.Padding = new Padding(0, 100, 0, 0);

            button.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(button);
            button.Click += CreateFloors;

            // Обертка для настроек пола (Правая сторона)
            settingsWrapper.AutoSize = false;
            settingsWrapper.Width = 800;
            settingsWrapper.Height = 400;
            settingsWrapper.AutoScroll = true;
            settingsWrapper.FlowDirection = FlowDirection.LeftToRight;
            settingsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(settingsWrapper);
            settingsWrapper.Anchor = AnchorStyles.Left;


            ICollection<Element>  rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
            foreach (Element room in rooms) { 
                Autodesk.Revit.DB.Parameter name = room.LookupParameter("Имя");
                if (name == null) name = room.LookupParameter("Name");
                if (name == null) continue;
                if (roomTypes.Contains(name.AsString())) continue;
                roomTypes.Add(name.AsString());
                
            }
            roomTypes = roomTypes.OrderBy(x => x).ToList();
            int i = 0;
            foreach (string roomType in roomTypes) {
                i++;
                FlowLayoutPanel newRoomType = AddRoomTypeSettings(roomType);
                if (i % 2 == 0) newRoomType.BackColor = System.Drawing.Color.FromArgb(255, 243, 221, 255);
            }
        }

        private void CreateFloors(object sender, EventArgs e)
        {


            ICollection<Element> rooms;
            switch (selectionCB.SelectedIndex)
            {
                case (1):
                    Autodesk.Revit.DB.View view = Doc.ActiveView;
                    rooms = new FilteredElementCollector(Doc, view.Id).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
                    break;

                case (2):                    
                    rooms = new FilteredElementCollector(Doc, SelectionIds).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
                    break;
                default:
                     rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();
                    break;
            }

            
            Transaction t = new Transaction(Doc, "Создать " + Mode.Name);
            t.Start();

            foreach (Element room in rooms)
            {
                SpatialElement roomSE = room as SpatialElement;
                SpatialElementBoundaryOptions boundaryOptions = new SpatialElementBoundaryOptions();
                CurveArray curveArray = new CurveArray();

                IList<CurveLoop> curveLoops = ExpIFCUtils.GetRoomBoundaryAsCurveLoopArray(roomSE, boundaryOptions, true);
                if ((curveLoops == null) || (curveLoops.Count == 0)) continue; 
                foreach (Curve curve in curveLoops[0]) {
                    curveArray.Append(curve);
                }

                Parameter roomName = room.LookupParameter("Имя");
                if (roomName == null) roomName = room.LookupParameter("Name");
                if (roomName != null) {
                    int i = roomTypes.IndexOf(roomName.AsString());
                    WinForms.ComboBox tempFloorTypesCB = (WinForms.ComboBox)settingsWrapper.Controls[i].Controls[1];
                    int j = tempFloorTypesCB.SelectedIndex;
                    int index = (j >= floorTypes.Count)?floorTypesCB.SelectedIndex:j;

                    FloorType floorType = FloorTypes[index] as FloorType;
                    Level level = Doc.GetElement(room.LevelId) as Level;
#if DEBUG2021 || REVIT2021
                    Floor newFloor = Doc.Create.NewFloor(curveArray, floorType, level, false, XYZ.BasisZ);

                    Parameter floorOffsetParam = newFloor.LookupParameter("Смещение от уровня");
                    if (floorOffsetParam == null) floorOffsetParam = newFloor.LookupParameter("Height Offset From Level");


                    if (floorOffsetParam != null)
                    {
                        WinForms.CheckBox offsetCB = (WinForms.CheckBox)settingsWrapper.Controls[i].Controls[3];
                        WinForms.NumericUpDown offsetNU = (WinForms.NumericUpDown)settingsWrapper.Controls[i].Controls[2];
                        Double value = 0;
                        if (offsetCB.Checked)
                        {
                            value = Convert.ToDouble(offsetTB.Value);                            
                        }
                        else 
                        {
                            value = Convert.ToDouble(offsetNU.Value);
                        }
                        floorOffsetParam.SetValueString(value.ToString());
                    }
#elif DEBUG2024 || REVIT2024

#endif
                }
            }

            t.Commit();

            this.DialogResult = DialogResult.OK;
            this.Close();
                
        }

        private FlowLayoutPanel AddRoomTypeSettings(string roomType) { 
            FlowLayoutPanel roomTypeSettings = new FlowLayoutPanel();
            roomTypeSettings.AutoSize = true;
            roomTypeSettings.FlowDirection = FlowDirection.LeftToRight;
            roomTypeSettings.Parent = settingsWrapper;
            settingsWrapper.Controls.Add(roomTypeSettings);
            settingsWrapper.SetFlowBreak(roomTypeSettings, true);
            roomTypeSettings.Anchor = AnchorStyles.Top;
            roomTypeSettings.Padding = new Padding(5, 1, 0, 1);

            // [0] Наименование типа комнат
            // [1] Тип пола
            // [2] Галочка "оффсет по умолчанию"
            // [3] Оффсет

            // [0] Наименование типа комнат
            WinForms.Label roomTypeLabel = new WinForms.Label();
            roomTypeLabel.Text = roomType;
            int lines = roomType.Length / 32 + 1;
            roomTypeLabel.Size = new Size(200, 15*lines);

            roomTypeLabel.Parent = roomTypeSettings;
            roomTypeSettings.Controls.Add(roomTypeLabel);
            roomTypeLabel.Anchor = AnchorStyles.Left;

            // [1] Тип пола
            WinForms.ComboBox roomTypeComboBox = new WinForms.ComboBox();
            foreach (Element floorType in FloorTypes)
            {
                roomTypeComboBox.Items.Add(floorType.Name);
            }
            roomTypeComboBox.Items.Add("По умолчанию");
            roomTypeComboBox.SelectedIndex = floorTypesCB.Items.Count;



            roomTypeComboBox.Parent = roomTypeSettings;
            roomTypeSettings.Controls.Add(roomTypeComboBox);
            roomTypeComboBox.Anchor = AnchorStyles.Left;

            roomTypeComboBox.Size = new Size(200, 20);

            // [2] Оффсет
            WinForms.NumericUpDown offsetRT = new WinForms.NumericUpDown();
            offsetRT.Increment = 1;
            offsetRT.Maximum = 100000;
            offsetRT.Minimum = -100000;

            offsetRT.Parent = roomTypeSettings;
            roomTypeSettings.Controls.Add(offsetRT);
            offsetRT.Anchor = AnchorStyles.Left;
            offsetRT.Size = new Size(50, 20);

            // [3] Галочка "оффсет по умолчанию"
            WinForms.CheckBox useDefault = new WinForms.CheckBox();
            useDefault.Text = "Отступ по умолчанию";
            useDefault.Checked = true;
            useDefault.Parent = roomTypeSettings;
            roomTypeSettings.Controls.Add(useDefault);
            useDefault.Anchor = AnchorStyles.Left;
            useDefault.Size = new Size(180, 20);

            return roomTypeSettings;
        }

        private class FunctionMode { 
            public BuiltInCategory BuiltInCategory;
            public string Name;
            public FunctionMode (BuiltInCategory builtInCategory, string name)
            {
                BuiltInCategory = builtInCategory;
                Name = name;                
            }        
        }
    }
}
