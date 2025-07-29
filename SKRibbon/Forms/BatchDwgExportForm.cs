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
using SKRibbon;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using System.IO;
using static SKRibbon.FormDesign;
using System.Text.RegularExpressions;

namespace SKRibbon
{
    public partial class BatchDwgExportForm : VForm
    {
        Document Doc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>> buildingsDict = new SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>>();
        string SavePath;
        VTextBox NameTextBox = new VTextBox();


        CheckBox cropRegionCheckBox = new CheckBox();
        System.Windows.Forms.ComboBox colorModeSelection = new System.Windows.Forms.ComboBox();
        Label pathLabel = new Label();
        FlowLayoutPanel optionsWrapper = new FlowLayoutPanel();


        public BatchDwgExportForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            SavePath = Path.GetDirectoryName(doc.PathName);
            if (SavePath == "")
            {
                SavePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            SavePath = Path.Combine(SavePath, "dwg");
            if (!Directory.Exists(SavePath))
            {
                Directory.CreateDirectory(SavePath);
            }
            if (SKRibbon.Properties.appSettings.Default.printFolder.Length == 0)
            {
                SKRibbon.Properties.appSettings.Default.printFolder = SavePath;
                SKRibbon.Properties.appSettings.Default.Save();
            }
            else
            {
                SavePath = SKRibbon.Properties.appSettings.Default.printFolder;
            }
            this.AutoScroll = true;
            this.Width = 710;
            this.Height = 480;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.BackColor = System.Drawing.Color.White;
            this.Text = "";
            this.ShowIcon = false;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Создаем дерево листов проекта
            buildingsDict = SKRibbon.FormUtils.CollectSheetDictionary(Doc, true);
            TreeView tree = SKRibbon.FormUtils.CreateSheetTreeView(buildingsDict);

            // Инициализируем параметры дерева и добавляем его в форму
            tree.CheckBoxes = true;
            tree.AfterCheck += node_AfterCheck;
            tree.Width = 300;
            tree.Height = 400;
            tree.Margin = new Padding(20, 10, 0, 0);
            tree.Parent = formWrapper;
            formWrapper.Controls.Add(tree);
            tree.Anchor = AnchorStyles.Left;

            // Добавляем wrapper для опций
            
            optionsWrapper.AutoSize = true;
            optionsWrapper.Anchor = AnchorStyles.Top;
            optionsWrapper.FlowDirection = FlowDirection.TopDown;
            optionsWrapper.Margin = new Padding(30, 10, 0, 0);
            optionsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(optionsWrapper);

            // Заголовок опций
            VHeaderLabel optionsHeader = new VHeaderLabel();
            optionsHeader.Text = "НАСТРОЙКИ DWG";
            optionsHeader.Size = new Size(300, 50);
            optionsHeader.Padding = new Padding(0, 0, 0, 0);

            optionsHeader.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(optionsHeader);

            int optionsWidth = 300;
            int fieldWidth = 210;

            // Опция 1. Путь сохранения файла
            // Заголовок
            Label pathHeader = new Label();
            pathHeader.Size = new Size(optionsWidth, 20);
            pathHeader.Margin = new Padding(0, 5, 0, 0);
            pathHeader.Text = "Файлы будут сохранены в папку:";
            pathHeader.Parent = optionsWrapper;
            pathHeader.Font = new Font(Label.DefaultFont, FontStyle.Bold);
            optionsWrapper.Controls.Add(pathHeader);

            // Добавляем путь
            pathLabel.Size = new Size(optionsWidth, (SavePath.Length / 44 + 1) * 15);
            pathLabel.Margin = new Padding(0, 0, 0, 0);
            pathLabel.Text = SavePath;            
            pathLabel.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(pathLabel);

            // Добавляем кнопку пути
            VButton pathButton = new VButton();
            pathButton.Text = "Выбрать другую папку";
            pathButton.Size = new Size(optionsWidth, 30);
            pathButton.Margin = new Padding(0, 10, 0, 30);
            pathButton.Click += ChooseFolder;
            pathButton.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(pathButton);
            pathButton.Anchor = AnchorStyles.Left;

            // Опция 2. Имя файла
            // Обертка
            FlowLayoutPanel lineWrapper_1 = new FlowLayoutPanel();
            lineWrapper_1.AutoSize = true;
            lineWrapper_1.FlowDirection = FlowDirection.LeftToRight;

            // Заголовок
            Label nameHeader = new Label();
            nameHeader.Size = new Size(optionsWidth - fieldWidth, 30);
            nameHeader.Margin = new Padding(0, 30, 0, 0);
            nameHeader.Text = "Имя файла:";
            nameHeader.Font = new Font(Label.DefaultFont, FontStyle.Bold);

            // Поле
            NameTextBox.Text = Doc.Title;
            NameTextBox.Size = new Size(fieldWidth, 50);
            NameTextBox.Margin = new Padding(0, 30, 0, 5);

            nameHeader.Parent = lineWrapper_1;
            lineWrapper_1.Controls.Add(nameHeader);

            NameTextBox.Parent = lineWrapper_1;
            lineWrapper_1.Controls.Add(NameTextBox);

            lineWrapper_1.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(lineWrapper_1);

            // Добавляем разделитель
            Label divider = new Label();
            divider.Text = "";
            divider.AutoSize = false;
            divider.Height = 2;
            divider.Width = 300;
            divider.BorderStyle = BorderStyle.Fixed3D;
            divider.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(divider);
            divider.Anchor = AnchorStyles.Top;

            // Добавляем кнопку запуска программы
            VButton okButton = new VButton();
            okButton.Text = "ВЫВЕСТИ ЛИСТЫ";
            okButton.Size = new Size(300, 60);
            okButton.Click += ExportSheets;
            okButton.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(okButton);
            okButton.Anchor = AnchorStyles.Bottom;
        }


        // Проставление галочек напротив всех "детей"
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    // Если у детей нода тоже есть дети, рекурсивно вызываем для них функцию
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        private void node_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }

        //Выбор папки
        public void ChooseFolder(object sender, EventArgs e)
        {       

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                pathLabel.Text = dialog.SelectedPath;
                SavePath = dialog.SelectedPath;
                SKRibbon.Properties.appSettings.Default.printFolder = SavePath;
                SKRibbon.Properties.appSettings.Default.Save();
            }
            pathLabel.Size = new Size(300, (SavePath.Length / 44 + 1) * 15);
        }

        private void ExportSheets(object sender, EventArgs e)
        {
            TreeView tree = (TreeView)formWrapper.Controls[0];
            
            foreach (TreeNode building in tree.Nodes)
            {                
                foreach (TreeNode tome in building.Nodes)
                {
                    List<ElementId> elemIds = new List<ElementId>();
                    string nameSuffix = "_";
                    if (!building.Text.Contains("ПРИМЕЧАНИЕ НЕ ЗАДАНО")) nameSuffix += building.Text + "_";
                    if (!tome.Text.Contains("РАЗДЕЛ НЕ ЗАДАН")) nameSuffix += tome.Text + "_";
                    nameSuffix = Regex.Replace(nameSuffix, @"\p{C}+", string.Empty); ;
                    nameSuffix = Regex.Replace(nameSuffix, @"[\~#%&*{}/:<>?|"",;']", string.Empty);
                    foreach (SKRibbon.FormUtils.SheetTreeNode sheet in tome.Nodes)
                    {
                        if (!sheet.Checked)
                        {
                            continue;
                        }
                        elemIds.Add(sheet.sheet.Id);
                    }
                    DWGExportOptions exportOptions = new DWGExportOptions();
                    exportOptions.MergedViews = true;
                    exportOptions.Colors = ExportColorMode.TrueColorPerView;

                    if (elemIds.Count != 0) Doc.Export(SavePath, NameTextBox.Text + nameSuffix, elemIds, exportOptions);
                }
            }
             //---------------------------------------------------------

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
