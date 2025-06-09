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
using SK_FD = SKRibbon.FormDesign;
using static SKRibbon.FormDesign;
using System.Windows.Controls;
namespace FillStamps
{
    [Transaction(TransactionMode.Manual)]
    public partial class FillStampsForm : VForm
    {
        Document Doc;
        FlowLayoutPanel formWrapper;
        FlowLayoutPanel linesWrapper;
        WinForms.CheckBox checkBox;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();

        public FillStampsForm(Document doc)
        {
            InitializeComponent();

            this.Text = "";
            this.BackColor = System.Drawing.Color.White;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle;
            this.Width = 880;
            this.Height = 460;


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
            WinForms.TreeView tree = SK_FU.CreateSheetTreeView(buildingsDict);

            // Инициализируем параметры дерева и добавляем его в форму
            tree.CheckBoxes = true;
            tree.AfterCheck += node_AfterCheck;
            tree.Width = 300;
            tree.Height = 365;
            tree.Margin = new Padding(15, 5, 0, 0);
            tree.Anchor = AnchorStyles.Left;

            //-----------------------------------------------------------------
            // Заголовок
            VHeaderLabel headerLabel = new VHeaderLabel();
            headerLabel.Size = new Size(600, 50);
            headerLabel.Text = "ДАННЫЕ О РАЗРАБОТЧИКАХ";
            headerLabel.Parent = linesWrapper;
            linesWrapper.Controls.Add(headerLabel);

            // Инициализация шаблона штампа

            for (int i = 1; i <= 6; i++)
            {
                FlowLayoutPanel lineWrapper = new FlowLayoutPanel();
                lineWrapper.FlowDirection = FlowDirection.LeftToRight;
                lineWrapper.AutoSize = true;

                WinForms.Label posLabel = new WinForms.Label();
                SK_FD.VTextBox posText = new SK_FD.VTextBox();
                WinForms.Label nameLabel = new WinForms.Label();
                SK_FD.VTextBox nameText = new SK_FD.VTextBox();

                int labelWidth = 100;
                int textBoxWidth = 150;
                int height = 20;

                //posLabel.BorderStyle = BorderStyle.FixedSingle;
                //nameLabel.BorderStyle = BorderStyle.FixedSingle;

                posLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
                nameLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);

                posLabel.Padding = new Padding(10, 0, 0, 0);

                posLabel.Size = new System.Drawing.Size(labelWidth, height);
                posText.Size = new System.Drawing.Size(textBoxWidth, height);
                nameLabel.Size = new System.Drawing.Size(labelWidth, height);
                nameText.Size = new System.Drawing.Size(textBoxWidth, height);

                posLabel.Anchor = AnchorStyles.Left;
                posText.Anchor = AnchorStyles.Left;
                nameLabel.Anchor = AnchorStyles.Left;
                nameText.Anchor = AnchorStyles.Left;

                posLabel.Text = "Должность " + i.ToString();
                nameLabel.Text = "Фамилия " + i.ToString();
                string posStr = "";
                switch (i)
                {
                    case 1:
                        posStr = "Разработал";
                        break;
                    case 2:
                        posStr = "Проверил";
                        break;
                    case 3: 
                        break;
                    case 4:
                        posStr = "Н.контр.";
                        break;
                    case 5:
                        posStr = "ГИП";
                        break;
                    case 6:
                        posStr = "ГАП";
                        break;
                }
                posText.Text = posStr;

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
            checkBox.Size = new System.Drawing.Size (600, 60);
            checkBox.Anchor = AnchorStyles.Top;
            checkBox.Padding = new Padding(10, 15, 0, 0);
            checkBox.Text = "Учитывать (перезаписывать) пустые поля";

            SK_FD.VButton button = new SK_FD.VButton();
            button.Size = new System.Drawing.Size(200, 50);
            button.Anchor = AnchorStyles.Top;
            button.Text = "ЗАПОЛНИТЬ ШТАМПЫ";
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
            WinForms.TreeView tree = (WinForms.TreeView)formWrapper.Controls[0];

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

                            FlowLayoutPanel lineWrapper = (FlowLayoutPanel)linesWrapper.Controls[i];
                            SK_FD.VTextBox tb1 = (SK_FD.VTextBox)lineWrapper.Controls[1];
                            SK_FD.VTextBox tb2 = (SK_FD.VTextBox)lineWrapper.Controls[3];

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
