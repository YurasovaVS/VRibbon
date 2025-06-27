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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Security.AccessControl;
using System.IO;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    public partial class AddSigForm : VForm
    {
        Document Doc;
        string Path;
        SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>> buildingsDict = new SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>>();

        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        TreeView sheetTree;
        Label newPath = new Label();
        public AddSigForm(Document doc, string path)
        {
            InitializeComponent();
            Doc = doc;
            Path = path;
                 
            //Создаем Wrapper для содержимого формы            
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.Padding = new Padding(5, 5, 5, 5);

            //Создаем заголовок
            Label header = new Label();
            header.Parent = formWrapper;
            formWrapper.Controls.Add(header);
            header.Anchor = AnchorStyles.Top;
            header.Size = new Size(500, 30);
            header.Text = "Путь к папке, где лежат DWG-подписи:";
            header.Font = new Font(Label.DefaultFont, FontStyle.Bold);

            //Создаем текстовое поле для пути
            newPath.Name = "pathLabel";
            newPath.Parent = formWrapper;
            formWrapper.Controls.Add(newPath);
            newPath.Size = new Size(500, 30);
            newPath.Text = Path;

            // Кнопка смены пути
            VButton pathButton = new VButton();
            pathButton.Anchor = AnchorStyles.Top;
            pathButton.Size = new Size(300, 30);
            pathButton.Text = "Сменить папку с подписями";
            pathButton.Click += ChooseFolder;

            pathButton.Parent = formWrapper;
            formWrapper.Controls.Add(pathButton);

            //Создаем заголовок чеклиста
            Label checkHeader = new Label();
            checkHeader.Parent = formWrapper;
            formWrapper.Controls.Add(checkHeader);
            checkHeader.Anchor = AnchorStyles.Top;
            checkHeader.Size = new Size(500, 30);
            checkHeader.Text = "Выберите листы:";
            checkHeader.Font = new Font(Label.DefaultFont, FontStyle.Bold);

            // Создаем словарь листов
            buildingsDict = FormUtils.CollectSheetDictionary(doc, true);
            sheetTree = FormUtils.CreateSheetTreeView(buildingsDict);

            sheetTree.Parent = formWrapper;
            formWrapper.Controls.Add(sheetTree);
            sheetTree.Anchor = AnchorStyles.Top;
            sheetTree.MinimumSize = new Size(500, 30);
            sheetTree.Height = 200;
            sheetTree.CheckBoxes = true;
            sheetTree.AfterCheck += node_AfterCheck;

            //Добавляем кнопку
            VButton button = new VButton();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Size = new Size (300, 60);

            button.Text = "ПРОСТАВИТЬ ПОДПИСИ";
            button.Click += PlaceSignatures;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;
        }

        public void PlaceSignatures(object sender, EventArgs e)
        {
            VButton button = (VButton)sender;
            
            string path = newPath.Text;
            StringBuilder sb = new StringBuilder();

            Transaction t = new Transaction(Doc, "Вставить подписи");
            t.Start();

            foreach (TreeNode building in sheetTree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (FormUtils.SheetTreeNode sheetNode in tome.Nodes)
                    {
                        if (!sheetNode.Checked)
                        {
                            continue;
                        }

                        ViewSheet sheet = sheetNode.sheet;
                        DWGImportOptions importOptions = new DWGImportOptions();

                        Double sheetRBpointX = sheet.Outline.Max.U;
                        Double sheetRBpointY = sheet.Outline.Min.V;

                        for (int i = 0; i <= 5; i++)
                        {
                            string paramName = "ADSK_Штамп Строка " + (i + 1).ToString() + " фамилия";
                            string paramValue = sheet.LookupParameter(paramName).AsString();

                            // Проверяем, задана ли фамилия
                            if (paramValue == null)
                            {
                                continue;
                            }
                            if (paramValue.Length == 0)
                            {
                                continue;
                            }

                            string signaturePath = path + "\\подпись_" + paramValue + ".dwg";
                            if (File.Exists(signaturePath))
                            {
                                ElementId signatureId;
                                Doc.Link(signaturePath, importOptions, sheet, out signatureId);
                                Element signature = Doc.GetElement(signatureId);
                                signature.Pinned = false;
                                Double x = sheetRBpointX - 0.49;
                                Double y = 30 - i * 5;
                                y = UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters);
                                y = sheetRBpointY + y;
                                XYZ move = new XYZ(x, y, 0);
                                signature.Location.Move(move);
                            }
                            else
                            {
                                sb.AppendLine("Подпись не найдена: " + signaturePath);
                            }
                        } // Конец перебора строк штампа
                    } // Конец foreach (sheet in tome)
                } // Конец foreach (tome in building)
            } // Конец foreach (building in tree)

            if (sb.Length == 0)
            {
                sb.AppendLine("Подписи проставлены. Ошибок нет.");
            }
            TaskDialog.Show("Ошибки", sb.ToString());

            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public void CheckAllChildNodes(TreeNode node, bool nodeChecker)
        {
            foreach (TreeNode childNode in node.Nodes)
            {
                childNode.Checked = nodeChecker;
                if (childNode.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(childNode, nodeChecker);
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

        public void ChooseFolder(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel wrapper = (FlowLayoutPanel)button.Parent;
            Label displayPath = (Label)wrapper.Controls[1];

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                displayPath.Text = dialog.SelectedPath;
                Path = dialog.SelectedPath;
            }
        }
    }

    //Класс для чеклиста листов
    public class CheckedSheetList : CheckedListBox
    {
        public List<Autodesk.Revit.DB.ViewSheet> sheetCollection;
        public CheckedSheetList()
        {
            sheetCollection = new List<Autodesk.Revit.DB.ViewSheet>();
        }
    }

    // Класс для нодов с листами
    public class SheetNode : TreeNode
    {
        public ViewSheet sheet;
    }
}
