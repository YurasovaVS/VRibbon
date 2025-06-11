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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using static SKRibbon.FormDesign;

namespace SKRibbon.Forms
{
    public partial class DeleteSigForm : VForm
    {
        Document Doc;
        SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>> buildingsDict = new SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>>();
        public DeleteSigForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;

            // Создаем Wrapper для содержимого формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.Padding = new Padding(5, 5, 5, 5);

            int panelWidth = 350;
            // Создаем заголовок
            Label header = new Label();
            header.Parent = formWrapper;
            formWrapper.Controls.Add(header);
            header.Anchor = AnchorStyles.Top;
            header.Size = new Size(panelWidth, 30);
            header.Text = "Выберите листы, с которых хотите удалить подписи:";
            header.Font = new Font(Label.DefaultFont, FontStyle.Bold);

            // Добавляем древо листов
            buildingsDict = FormUtils.CollectSheetDictionary(doc, true);
            TreeView sheetTree = FormUtils.CreateSheetTreeView(buildingsDict);

            sheetTree.Parent = formWrapper;
            formWrapper.Controls.Add(sheetTree);
            sheetTree.Anchor = AnchorStyles.Top;
            sheetTree.MinimumSize = new Size(panelWidth, 30);
            sheetTree.Height = 200;
            sheetTree.CheckBoxes = true;
            sheetTree.AfterCheck += node_AfterCheck;

            // Добавляем кнопку ОК
            VButton button = new VButton();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Size = new Size (panelWidth, 50);

            button.Text = "Убрать подписи";
            button.Click += RemoveSignatures;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;
        }

        public void RemoveSignatures (object sender, EventArgs e)
        {
            VButton button = (VButton)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;
            TreeView sheetTree = (TreeView)formWrapper.Controls[1];

            Transaction t = new Transaction(Doc, "Убрать подписи");
            t.Start();

            foreach (TreeNode building in sheetTree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (FormUtils.SheetTreeNode sheetNode in tome.Nodes)
                    {
                        if (sheetNode.Checked)
                        {
                            ViewSheet sheet = sheetNode.sheet;
                            ICollection<Element> signatures = new FilteredElementCollector(Doc, sheet.Id).
                                                            OfClass(typeof(ImportInstance)).
                                                            ToElements();
                            foreach (ImportInstance signature in signatures)
                            {
                                string name = signature.LookupParameter("Имя").AsString();
                                string pattern = "подпись_*";
                                if (Regex.Match(name, pattern).Success)
                                {
                                    Doc.Delete(signature.Id);
                                }
                            }
                        }
                    }
                }
            }

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
    }
    public class CheckedSheetList : CheckedListBox
    {
        public List<Autodesk.Revit.DB.ViewSheet> sheetCollection;
        public CheckedSheetList()
        {
            sheetCollection = new List<Autodesk.Revit.DB.ViewSheet>();
        }
    }
}
