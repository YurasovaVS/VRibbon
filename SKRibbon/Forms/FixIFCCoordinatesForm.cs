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
using System.Text.RegularExpressions;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;


namespace SKRibbon
{
    public partial class FixIFCCoordinatesForm: WinForms.Form
    {
        Document Doc;
        String IFCFilePath = "C:\\Users\\yuras\\Documents\\!Временное\\_2025-04-22_Revit_Icons\\test.ifc";

        Dictionary<string, Element> IFC_Links = new Dictionary<string, Element>();

        int LabelWidth = 150;
        int TextBoxWidth = 300;
        int RowHeight = 20;

        WinForms.FlowLayoutPanel FormWrapper = new WinForms.FlowLayoutPanel();

        WinForms.ComboBox FilePathTextBox = new WinForms.ComboBox();

        WinForms.TextBox CoordinateX_TextBox = new WinForms.TextBox();
        WinForms.TextBox CoordinateY_TextBox = new WinForms.TextBox();
        WinForms.TextBox CoordinateZ_TextBox = new WinForms.TextBox();
        WinForms.TextBox CoordinateRotation_TextBox = new WinForms.TextBox();

        double Angle = 0.0;
        public FixIFCCoordinatesForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;

            this.AutoScroll = true;
            this.Width = 550;
            this.Height = 350;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle;
            this.Text = "Исправить IFC файл";

            FormWrapper.AutoSize = true;
            FormWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            FormWrapper.Parent = this;
            this.Controls.Add(FormWrapper);

            // Получаем координаты базовой точки файла
            Element projectInfoElement = new FilteredElementCollector(doc)
                                            .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                                            .FirstElement();

            Autodesk.Revit.DB.Parameter paramX = projectInfoElement.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM);
            Autodesk.Revit.DB.Parameter paramY = projectInfoElement.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM);
            Autodesk.Revit.DB.Parameter paramZ = projectInfoElement.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM);
            Autodesk.Revit.DB.Parameter paramAngle = projectInfoElement.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM);


            string tempString = "";

            WinForms.Label label1 = new WinForms.Label();
            WinForms.Label label2 = new WinForms.Label();

            label1.Size = label2.Size = new Size(LabelWidth + TextBoxWidth, RowHeight);
            label1.Font = label2.Font = new System.Drawing.Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            label1.TextAlign = label2.TextAlign = ContentAlignment.MiddleCenter;

            label1.Text = "IFC-файл";
            label2.Text = "Координаты";


            label1.Parent = FormWrapper;
            FormWrapper.Controls.Add(label1);

            WinForms.FlowLayoutPanel Panel_FilePath = AddLabelComboBoxPanel("Укажите путь:", FilePathTextBox, tempString);

            label2.Parent = FormWrapper;
            FormWrapper.Controls.Add(label2);

            double tempDouble = UnitUtils.ConvertFromInternalUnits(paramX.AsDouble(), UnitTypeId.Millimeters);
            WinForms.FlowLayoutPanel Panel_CoordinateX = AddLabelTextBoxPanel("Х:", CoordinateX_TextBox, tempDouble.ToString());

            tempDouble = UnitUtils.ConvertFromInternalUnits(paramY.AsDouble(), UnitTypeId.Millimeters);
            WinForms.FlowLayoutPanel Panel_CoordinateY = AddLabelTextBoxPanel("Y:", CoordinateY_TextBox, tempDouble.ToString());

            tempDouble = UnitUtils.ConvertFromInternalUnits(paramZ.AsDouble(), UnitTypeId.Millimeters);
            WinForms.FlowLayoutPanel Panel_CoordinateZ = AddLabelTextBoxPanel("Z:", CoordinateZ_TextBox, tempDouble.ToString());

            Angle = tempDouble = UnitUtils.ConvertFromInternalUnits(paramAngle.AsDouble(), UnitTypeId.Degrees); 
            WinForms.FlowLayoutPanel Panel_CoordinateRotation = AddLabelTextBoxPanel("Угол поворота:", CoordinateRotation_TextBox, tempDouble.ToString());

            WinForms.Button okButton = new WinForms.Button();
            okButton.Text = "Исправить файл";
            okButton.Size = new Size(LabelWidth + TextBoxWidth, RowHeight);
            okButton.Click += RunFixingIFC;
            okButton.Parent = FormWrapper;
            FormWrapper.Controls.Add(okButton);


            // Находим все ссылки в проекте
            //ICollection<ElementId> linkRefs= ExternalFileUtils.GetAllExternalFileReferences(doc);

            ICollection<Element> links = new FilteredElementCollector(Doc).
                                                OfClass(typeof(RevitLinkType)).
                                                ToElements();
            // Заполняем comboBox с названиями IFC ссылок
            foreach (Element link in links)
            {
                if (link.Name.EndsWith(".ifc"))
                {
                    FilePathTextBox.Items.Add(link.Name);
                    IFC_Links.Add(link.Name, link);
                }
            }
            // Если ссылок нет, блокируем кнопку
            if (FilePathTextBox.Items.Count == 0)
            {
                okButton.Enabled = false;
                okButton.Text = "В файле нет ссылок на IFC";
                FilePathTextBox.Enabled = false;
            }
            else
            {
                FilePathTextBox.SelectedIndex = 0;
            }
        }

        private void RunFixingIFC (object sender, EventArgs e)
        {
            Element linkElement = IFC_Links[FilePathTextBox.Text];
            RevitLinkType link = linkElement as RevitLinkType;

            ExternalFileReference exFileRef = link.GetExternalFileReference();
            ModelPath modelPath = exFileRef.GetAbsolutePath();
            IFCFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
            IFCFilePath = IFCFilePath.Substring(0, IFCFilePath.Length - 4);
            string line555555 = "#555555=IFCCARTESIANPOINT";
            string line666666 = "#666666=IFCDIRECTION";
            string line777777 = "#777777=IFCAXIS2PLACEMENT3D";

            string line555555_add = "((" + FormattedCoordinate(CoordinateX_TextBox.Text) + ',' 
                                         + FormattedCoordinate(CoordinateY_TextBox.Text) + ',' 
                                         + FormattedCoordinate(CoordinateZ_TextBox.Text) + "));";

            string cos = Math.Cos(Angle * Math.PI / 180).ToString();
            string sin = Math.Sin((360 - Angle) * Math.PI / 180).ToString();

            string line666666_add = "((" + FormattedCoordinate(cos) + ',' + FormattedCoordinate(sin) + ',' + "0." + "));";
            string line777777_add = "(#555555,#4,#666666);";

            string searchPattern = @"IFCLOCALPLACEMENT\(\$,#\d+\);";
            string replacement = @"IFCLOCALPLACEMENT($,#777777);";

            link.Unload(null);

            try
            {
                // Считывваем все строки IFC файла
                string[] lines = File.ReadAllLines(IFCFilePath);

                int firstLineIndex = 0;

                bool flag555555 = false;
                bool flag666666 = false;
                bool flag777777 = false;
                
                for (int i = 0; i < lines.Length; i++)
                {
                    // Если есть строки 555555, 666666, 777777 - перезаписываем их
                    if (lines[i].Contains(line555555))
                    {
                        lines[i] = line555555 + line555555_add;
                        flag555555 = true;
                    }
                    if (lines[i].Contains(line666666)) 
                    { 
                        lines[i] = line666666 + line666666_add;
                        flag666666 = true;
                    }
                    if (lines[i].Contains(line777777))
                    {
                        lines[i] = line777777 + line777777_add;
                        flag777777 = true;
                    }
                    if (lines[i].StartsWith("#1=")) firstLineIndex = i + 1;

                    if (lines[i].Contains("IFCLOCALPLACEMENT($,"))
                    {
                        // Use regex to replace the pattern
                        lines[i] = Regex.Replace(
                            lines[i],
                            @"IFCLOCALPLACEMENT\(\$,#\d+\);",
                            replacement
                        );
                    }
                }

                List<string> linesList = lines.ToList();

                // Если строк нет, вставляем их после строки #1
                if (!flag777777) 
                {
                    linesList.Insert(firstLineIndex, line777777 + line777777_add);
                }

                if (!flag666666)
                {
                    linesList.Insert(firstLineIndex, line666666 + line666666_add);
                }

                if (!flag666666)
                {
                    linesList.Insert(firstLineIndex, line555555 + line555555_add);
                }

                lines = linesList.ToArray();

                // Перезаписываем файл
                File.Delete(IFCFilePath);
                File.WriteAllLines(IFCFilePath, lines);
                Console.WriteLine($"File processed successfully. Output: {IFCFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file: {ex.Message}");
            }
            // Перезагружаем IFC
            link.Reload();
            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
        }

        WinForms.FlowLayoutPanel AddLabelTextBoxPanel(string Name, WinForms.TextBox textBox, string text) {
            WinForms.FlowLayoutPanel panel = new WinForms.FlowLayoutPanel();
            panel.AutoSize = true;
            panel.FlowDirection = WinForms.FlowDirection.LeftToRight;

            WinForms.Label label = new WinForms.Label();
            label.Text = Name;
            label.Size = new Size(LabelWidth, RowHeight);

            textBox.Size = new Size(TextBoxWidth, RowHeight);
            textBox.Text = text;

            label.Parent = panel;
            panel.Controls.Add(label);

            textBox.Parent = panel;
            panel.Controls.Add(textBox);

            panel.Parent = FormWrapper;
            FormWrapper.Controls.Add(panel);

            return panel;
        }

        WinForms.FlowLayoutPanel AddLabelComboBoxPanel(string Name, WinForms.ComboBox comboBox, string text)
        {
            WinForms.FlowLayoutPanel panel = new WinForms.FlowLayoutPanel();
            panel.AutoSize = true;
            panel.FlowDirection = WinForms.FlowDirection.LeftToRight;

            WinForms.Label label = new WinForms.Label();
            label.Text = Name;
            label.Size = new Size(LabelWidth, RowHeight);

            comboBox.Size = new Size(TextBoxWidth, RowHeight);
            comboBox.Text = text;

            label.Parent = panel;
            panel.Controls.Add(label);

            comboBox.Parent = panel;
            panel.Controls.Add(comboBox);

            panel.Parent = FormWrapper;
            FormWrapper.Controls.Add(panel);

            return panel;
        }

        string FormattedCoordinate (string coord)
        {
            string fixedCoord = coord.Replace(',', '.');
            if (!fixedCoord.Contains('.')) fixedCoord += '.';
            return fixedCoord;
        }
    }
}
