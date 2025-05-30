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
using Autodesk.Revit.Attributes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using SK_FU = SKRibbon.FormUtils;
namespace FillStamps
{
    [Transaction(TransactionMode.Manual)]
    public partial class FillStampsForm : WinForms.Form
    {
        Document Doc;
        FlowLayoutPanel formWrapper;
        FlowLayoutPanel linesWrapper;
        WinForms.CheckBox checkBox;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();

        public FillStampsForm(Document doc)
        {
            InitializeComponent();

            this.Text = "Заполнить штампы";
            Doc = doc;

            formWrapper = new FlowLayoutPanel();
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.AutoSize = true;

            linesWrapper = new FlowLayoutPanel();
            linesWrapper.FlowDirection = FlowDirection.TopDown;
            linesWrapper.AutoSize = true;


            //Собираем все листы в основу для древа
            /* Словарь Зданий
             *      Здание : Словарь Томов
             *          Том : Список объектов
             *              Объект ViewSheet
             *                  
            */

            buildingsDict = SK_FU.CollectSheetDictionary(doc, true);
            TreeView tree = SK_FU.CreateSheetTreeView(buildingsDict);

            // Инициализируем параметры дерева и добавляем его в форму
            tree.CheckBoxes = true;
            tree.AfterCheck += node_AfterCheck;
            tree.Width = 300;
            tree.Height = 400;
            tree.Margin = new Padding(0, 10, 0, 0);
            tree.Anchor = AnchorStyles.Left;

            //-----------------------------------------------------------------
            // Инициализация шаблона штампа

            for (int i = 1; i <= 6; i++)
            {
                FlowLayoutPanel lineWrapper = new FlowLayoutPanel();
                lineWrapper.FlowDirection = FlowDirection.LeftToRight;
                lineWrapper.AutoSize = true;

                Label posLabel = new Label();
                WinForms.TextBox posText = new WinForms.TextBox();
                Label nameLabel = new Label();
                WinForms.TextBox nameText = new WinForms.TextBox();

                posLabel.Size = new System.Drawing.Size(100, 30);
                posText.Size = new System.Drawing.Size(100, 30);
                nameLabel.Size = new System.Drawing.Size(100, 30);
                nameText.Size = new System.Drawing.Size(100, 30);

                posLabel.Anchor = AnchorStyles.Left;
                posText.Anchor = AnchorStyles.Left;
                nameLabel.Anchor = AnchorStyles.Left;
                nameText.Anchor = AnchorStyles.Left;

                posLabel.Text = "Должность " + i.ToString();
                nameLabel.Text = "Фамилия " + i.ToString();

                lineWrapper.Controls.Add(posLabel);
                lineWrapper.Controls.Add(posText);
                lineWrapper.Controls.Add(nameLabel);
                lineWrapper.Controls.Add(nameText);

                posLabel.Parent = lineWrapper;
                posText.Parent = lineWrapper;
                nameLabel.Parent = lineWrapper;
                nameText.Parent = lineWrapper;

                lineWrapper.Parent = linesWrapper;
                linesWrapper.Controls.Add (lineWrapper);
            }

            checkBox = new WinForms.CheckBox();
            checkBox.Size = new System.Drawing.Size (200, 60);
            checkBox.Anchor = AnchorStyles.Top;
            checkBox.Text = "Учитывать (перезаписывать) пустые поля";

            Button button = new Button();
            button.Size = new System.Drawing.Size(150, 100);
            button.Anchor = AnchorStyles.Top;
            button.Text = "Заполнить штампы";
            button.Click += RunProgram;

            checkBox.Parent = linesWrapper;
            linesWrapper.Controls.Add(checkBox);

            button.Parent = linesWrapper;
            linesWrapper.Controls.Add(button);

            tree.Parent = formWrapper;
            formWrapper.Controls.Add(tree);

            linesWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(linesWrapper);

            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);
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

        private void RunProgram(object sender, EventArgs e)
        {
            TreeView tree = (TreeView)formWrapper.Controls[0];

            Transaction t = new Transaction(Doc, "Заполнить штампы");
            t.Start();

            foreach (TreeNode building in tree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (SK_FU.SheetTreeNode sheetNode in tome.Nodes)
                    {
                        if (!sheetNode.Checked)
                        {
                            continue;
                        }

                        ViewSheet sheet = sheetNode.sheet;

                        for (int i = 1; i<=6; i++)
                        {
                            string paramPos_name = "ADSK_Штамп Строка " + (i).ToString() + " должность";
                            Parameter paramPos = sheet.LookupParameter(paramPos_name);

                            // Проверяем, задана ли должность
                            if (paramPos == null)
                            {
                                continue;
                            }

                            string paramName_name = "ADSK_Штамп Строка " + (i).ToString() + " фамилия";
                            Parameter paramName = sheet.LookupParameter(paramName_name);

                            // Проверяем, задана ли фамилия
                            if (paramName == null)
                            {
                                continue;
                            }

                            FlowLayoutPanel lineWrapper = (FlowLayoutPanel)linesWrapper.Controls[i-1];
                            WinForms.TextBox tb1 = (WinForms.TextBox)lineWrapper.Controls[1];
                            WinForms.TextBox tb2 = (WinForms.TextBox)lineWrapper.Controls[3];

                            // ПРОВЕРИТЬ НА ПЕРЕБИВАНИЕ

                            if ((tb1.Text.Length > 0) || (checkBox.Checked)) paramPos.Set(tb1.Text);
                            if ((tb2.Text.Length > 0) || (checkBox.Checked)) paramName.Set(tb2.Text);
                        }
                    }
                }
            }
            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
