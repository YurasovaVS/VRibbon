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
using System.Windows.Media.Media3D;
using System.Xml;
namespace FillStamps
{
    [Transaction(TransactionMode.Manual)]
    public partial class FillStampsForm : VForm
    {
        Document Doc;
        FlowLayoutPanel formWrapper;
        FlowLayoutPanel linesWrapper;
        WinForms.CheckBox checkBox;
        SK_FD.VTextBox dateTextBox;
        WinForms.CheckBox dateCheckBox;
        SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>> buildingsDict = new SortedDictionary<string, SortedDictionary<string, List<ViewSheet>>>();


        FlowLayoutPanel advSettingsPanel = new FlowLayoutPanel();
        VHeaderLabel advSettingsLabel = new VHeaderLabel();

        SK_FD.VTextBox settingTexBox_1_1 = new SK_FD.VTextBox();
        SK_FD.VTextBox settingTexBox_1_2 = new SK_FD.VTextBox();
        SK_FD.VTextBox settingTexBox_2_1 = new SK_FD.VTextBox();
        SK_FD.VTextBox settingTexBox_2_2 = new SK_FD.VTextBox();

        WinForms.Label advSetHeader_1 = new WinForms.Label();
        WinForms.Label advSetHeader_2 = new WinForms.Label();

        WinForms.CheckBox advSettingsCheckBox;

        public FillStampsForm(Document doc)
        {
            InitializeComponent();

            this.Text = "";
            this.BackColor = System.Drawing.Color.White;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle;
            this.Width = 880;
            this.Height = 500;


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
            tree.Width = 360;
            tree.Height = 700;
            tree.Margin = new Padding(15, 5, 0, 0);
            tree.Anchor = AnchorStyles.Left;

            //-----------------------------------------------------------------
            // Опции
            // 1. Заголовок
            VHeaderLabel headerLabel = new VHeaderLabel();
            headerLabel.Size = new Size(600, 50);
            headerLabel.Text = "ДАННЫЕ О РАЗРАБОТЧИКАХ";
            headerLabel.Parent = linesWrapper;
            linesWrapper.Controls.Add(headerLabel);

            // 2. Инициализация шаблона штампа (6 строк)

            int labelWidth = 100;
            int textBoxWidth = 150;
            int height = 20;

            for (int i = 1; i <= 6; i++)
            {
                FlowLayoutPanel lineWrapper = new FlowLayoutPanel();
                lineWrapper.FlowDirection = FlowDirection.LeftToRight;
                lineWrapper.AutoSize = true;

                WinForms.Label posLabel = new WinForms.Label();
                SK_FD.VTextBox posText = new SK_FD.VTextBox();
                WinForms.Label nameLabel = new WinForms.Label();
                SK_FD.VTextBox nameText = new SK_FD.VTextBox();

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

            // 3. Галочка "Учитывать (перезаписывать) пустые поля"
            checkBox = new WinForms.CheckBox();
            checkBox.Size = new System.Drawing.Size (600, 60);
            checkBox.Anchor = AnchorStyles.Top;
            checkBox.Padding = new Padding(10, 15, 0, 0);
            checkBox.Text = "Перезаписывать поля, если они пустые";

            // Разделители
            VDivider divider1 = new VDivider();
            VDivider divider2 = new VDivider();
            VDivider divider3 = new VDivider();

            // 4. Дата принятия штампа
            // 4.1. Обертка даты
            FlowLayoutPanel dateLineWrapper = new FlowLayoutPanel();
            dateLineWrapper.FlowDirection = FlowDirection.LeftToRight;
            dateLineWrapper.AutoSize = true;
            dateLineWrapper.Margin = new Padding(0, 10, 0, 0);

            // 4.2. Лейбл даты
            WinForms.Label dateLabel = new WinForms.Label();
            dateLabel.Text = "Дата";
            dateLabel.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            dateLabel.Padding = new Padding(10, 0, 0, 0);
            dateLabel.Anchor = AnchorStyles.Left;

            // 4.3. Поле даты
            dateTextBox = new VTextBox();
            dateTextBox.Size = new System.Drawing.Size(textBoxWidth, height);
            dateTextBox.Anchor = AnchorStyles.Left;

            // Берем текущую дату
            DateTime today = DateTime.Now;
            dateTextBox.Text = today.ToString("MM.yy");


            dateLineWrapper.Controls.Add(dateLabel);
            dateLineWrapper.Controls.Add(dateTextBox);
            dateLabel.Parent = dateLineWrapper;
            dateTextBox.Parent = dateLineWrapper;

            // 5. Галочка перезаписывания даты            
            dateCheckBox = new WinForms.CheckBox();
            dateCheckBox.Size = new System.Drawing.Size(600, 60);
            dateCheckBox.Anchor = AnchorStyles.Top;
            dateCheckBox.Padding = new Padding(10, 0, 0, 0);
            dateCheckBox.Text = "Перезаписывать дату, если она пустая";
            dateCheckBox.Checked = true;

            // Дополнительные опции
            // Заголовок
            advSettingsLabel.Size = new Size(600, 40);
            advSettingsLabel.Text = "ПРОДВИНУТЫЕ НАСТРОЙКИ";
            advSettingsLabel.ForeColor = System.Drawing.Color.Gray;

            // 0. Обертка (не особо нужная, но мне так удобнее)
            advSettingsPanel.FlowDirection = FlowDirection.TopDown;
            advSettingsPanel.AutoSize = true;
            advSettingsPanel.ForeColor = System.Drawing.Color.Gray;

            advSettingsPanel.Controls.Add(advSettingsLabel);
            advSettingsLabel.Parent = advSettingsPanel;

            // 1. Галочка включения доп. опций
            advSettingsCheckBox = new WinForms.CheckBox();
            advSettingsCheckBox.Size = new System.Drawing.Size(600, 60);
            advSettingsCheckBox.Anchor = AnchorStyles.Top;
            advSettingsCheckBox.Padding = new Padding(10, 0, 0, 0);
            advSettingsCheckBox.Margin = new Padding(0, 0, 0, 0);
            advSettingsCheckBox.Text = "Я знаю, что делаю.";
            advSettingsCheckBox.Checked = false;
            advSettingsCheckBox.CheckedChanged += UnlockAdvancedSettings;

            advSettingsPanel.Controls.Add(advSettingsCheckBox);
            advSettingsCheckBox.Parent = advSettingsPanel;

            // Заголовок перед строкой
            advSetHeader_1.Text = "ПАРАМЕТР ДОЛЖНОСТИ";
            advSetHeader_1.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            advSetHeader_1.Anchor = AnchorStyles.Left;
            advSetHeader_1.Size = new Size (labelWidth * 2 + textBoxWidth * 2, height);
            advSetHeader_1.Margin = new Padding(5, 0, 0, 0);

            advSettingsPanel.Controls.Add(advSetHeader_1);
            advSetHeader_1.Parent = advSettingsPanel;

            // 2. Строка "Параметр должности"
            // 2.1. Обертка
            FlowLayoutPanel settingPanel_1 = new FlowLayoutPanel();
            settingPanel_1.FlowDirection = FlowDirection.LeftToRight;
            settingPanel_1.AutoSize = true;

            advSettingsPanel.Controls.Add(settingPanel_1);
            settingPanel_1.Parent = advSettingsPanel;

            // 2.2. Лейбл 1
            WinForms.Label settingLabel_1_1 = new WinForms.Label();
            settingLabel_1_1.Text = "Префикс";
            settingLabel_1_1.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            settingLabel_1_1.Anchor = AnchorStyles.Left;
            settingLabel_1_1.Size = new Size(labelWidth, height);
            

            settingPanel_1.Controls.Add(settingLabel_1_1);
            settingLabel_1_1.Parent = settingPanel_1;

            // 2.3. Текстбокс 1
            settingTexBox_1_1.Size = new Size(textBoxWidth, height);
            settingTexBox_1_1.Text = "ADSK_Штамп Строка ";
            settingTexBox_1_1.Enabled = false;
            settingTexBox_1_1.TextChanged += onAdvancedSettingChange;

            settingPanel_1.Controls.Add(settingTexBox_1_1);
            settingTexBox_1_1.Parent = settingPanel_1;

            // 2.4. Лейбл 2
            WinForms.Label settingLabel_1_2 = new WinForms.Label();
            settingLabel_1_2.Text = "Суффикс";
            settingLabel_1_2.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            settingLabel_1_2.Anchor = AnchorStyles.Left;
            settingLabel_1_2.Size = new Size(labelWidth, height);

            settingPanel_1.Controls.Add(settingLabel_1_2);
            settingLabel_1_2.Parent = settingPanel_1;

            // 2.5. Текстбокс 2
            settingTexBox_1_2.Size = new Size(textBoxWidth, height);
            settingTexBox_1_2.Text = " должность";
            settingTexBox_1_2.Enabled = false;
            settingTexBox_1_2.TextChanged += onAdvancedSettingChange;


            settingPanel_1.Controls.Add(settingTexBox_1_2);
            settingTexBox_1_2.Parent = settingPanel_1;

            // Заголовок перед строкой
            advSetHeader_2.Text = "ПАРАМЕТР ФАМИЛИИ";
            advSetHeader_2.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            advSetHeader_2.Anchor = AnchorStyles.Left;
            advSetHeader_2.Size = new Size(labelWidth * 2 + textBoxWidth * 2, height);
            advSetHeader_2.Margin = new Padding(5, 10, 0, 0);

            advSettingsPanel.Controls.Add(advSetHeader_2);
            advSetHeader_2.Parent = advSettingsPanel;

            // 3. Строка "Параметр фамилии"
            // 3.1. Обертка
            FlowLayoutPanel settingPanel_2 = new FlowLayoutPanel();
            settingPanel_2.FlowDirection = FlowDirection.LeftToRight;
            settingPanel_2.AutoSize = true;

            advSettingsPanel.Controls.Add(settingPanel_2);
            settingPanel_2.Parent = advSettingsPanel;

            // 3.2. Лейбл 1
            WinForms.Label settingLabel_2_1 = new WinForms.Label();
            settingLabel_2_1.Text = "Префикс";
            settingLabel_2_1.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            settingLabel_2_1.Anchor = AnchorStyles.Left;
            settingLabel_2_1.Size = new Size(labelWidth, height);

            settingPanel_2.Controls.Add(settingLabel_2_1);
            settingLabel_2_1.Parent = settingPanel_2;

            // 3.3. Текстбокс 1
            settingTexBox_2_1.Size = new Size(textBoxWidth, height);
            settingTexBox_2_1.Text = "ADSK_Штамп Строка ";
            settingTexBox_2_1.Enabled = false;
            settingTexBox_2_1.TextChanged += onAdvancedSettingChange;

            settingPanel_2.Controls.Add(settingTexBox_2_1);
            settingTexBox_2_1.Parent = settingPanel_2;


            // 3.4. Лейбл 2
            WinForms.Label settingLabel_2_2 = new WinForms.Label();
            settingLabel_2_2.Text = "Суффикс";
            settingLabel_2_2.Font = new Font(WinForms.Label.DefaultFont, FontStyle.Bold);
            settingLabel_2_2.Anchor = AnchorStyles.Left;
            settingLabel_2_2.Size = new Size(labelWidth, height);

            settingPanel_2.Controls.Add(settingLabel_2_2);
            settingLabel_2_2.Parent = settingPanel_2;

            // 3.5. Текстбокс 2
            settingTexBox_2_2.Size = new Size(textBoxWidth, height);
            settingTexBox_2_2.Text = " фамилия";
            settingTexBox_2_2.Enabled = false;
            settingTexBox_2_2.TextChanged += onAdvancedSettingChange;

            settingPanel_2.Controls.Add(settingTexBox_2_2);
            settingTexBox_2_2.Parent = settingPanel_2;

            advSetHeader_1.Text += " (" + settingTexBox_1_1.Text + "i" + settingTexBox_1_2.Text + ")";
            advSetHeader_2.Text += " (" + settingTexBox_2_1.Text + "i" + settingTexBox_2_2.Text + ")";

            // 4. Строка "Параметр даты"


            // Кнопка запуска программы

            SK_FD.VButton button = new SK_FD.VButton();
            button.Size = new System.Drawing.Size(labelWidth * 2 + textBoxWidth * 2, 50);
            button.Margin = new Padding(5);
            button.Anchor = AnchorStyles.Top;
            button.Text = "ЗАПОЛНИТЬ ШТАМПЫ";
            button.Click += RunProgram;

            // Порядок элементов в настройках

            checkBox.Parent = linesWrapper;
            linesWrapper.Controls.Add(checkBox);

            divider1.Parent = linesWrapper;
            linesWrapper.Controls.Add(divider1);

            dateLineWrapper.Parent = linesWrapper;
            linesWrapper.Controls.Add(dateLineWrapper);

            dateCheckBox.Parent = linesWrapper;
            linesWrapper.Controls.Add(dateCheckBox);

            // Кнопка
            button.Parent = linesWrapper;
            linesWrapper.Controls.Add(button);

            // Разделитель
            divider2.Margin = new Padding(0, 10, 0, 10);   
            divider2.Parent = linesWrapper;
            linesWrapper.Controls.Add(divider2);

            // Продвинутые настройки

            advSettingsPanel.Parent = linesWrapper;
            linesWrapper.Controls.Add(advSettingsPanel);

            //

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

        private void UnlockAdvancedSettings(object sender, EventArgs e) 
        {
            settingTexBox_1_1.Enabled = advSettingsCheckBox.Checked;
            settingTexBox_1_2.Enabled = advSettingsCheckBox.Checked;
            settingTexBox_2_1.Enabled = advSettingsCheckBox.Checked;
            settingTexBox_2_2.Enabled = advSettingsCheckBox.Checked;

            if (advSettingsCheckBox.Checked)
            {
                advSettingsPanel.ForeColor = System.Drawing.Color.Black;
                advSettingsLabel.ForeColor = System.Drawing.Color.FromArgb(255, 174, 112, 199);
            }
            else
            {
                advSettingsPanel.ForeColor = System.Drawing.Color.Gray;
                advSettingsLabel.ForeColor = System.Drawing.Color.Gray;
            }

        }

        private void onAdvancedSettingChange (object sender, EventArgs e)
        {
            advSetHeader_1.Text = "ПАРАМЕТР ДОЛЖНОСТИ (" + settingTexBox_1_1.Text + "i" + settingTexBox_1_2.Text + ")";
            advSetHeader_2.Text = "ПАРАМЕТР ФАМИЛИИ (" + settingTexBox_2_1.Text + "i" + settingTexBox_2_2.Text + ")";
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
                            string paramPos_name = settingTexBox_1_1.Text + (i).ToString() + settingTexBox_1_2.Text;
                            Parameter paramPos = sheet.LookupParameter(paramPos_name);

                            // Проверяем, задана ли должность
                            if (paramPos == null)
                            {
                                continue;
                            }

                            string paramName_name = settingTexBox_2_1.Text + (i).ToString() + settingTexBox_2_2.Text;
                            Parameter paramName = sheet.LookupParameter(paramName_name);

                            // Проверяем, задана ли фамилия
                            if (paramName == null)
                            {
                                continue;
                            }

                            FlowLayoutPanel lineWrapper = (FlowLayoutPanel)linesWrapper.Controls[i];
                            SK_FD.VTextBox tb1 = (SK_FD.VTextBox)lineWrapper.Controls[1];
                            SK_FD.VTextBox tb2 = (SK_FD.VTextBox)lineWrapper.Controls[3];

                            // Проверка на перебивание

                            if ((tb1.Text.Length > 0) || (checkBox.Checked)) paramPos.Set(tb1.Text);
                            if ((tb2.Text.Length > 0) || (checkBox.Checked)) paramName.Set(tb2.Text);
                        }

                        if ((dateTextBox.MaxLength > 0) || (dateCheckBox.Checked))
                        {
                            
                            Parameter dateParam = sheet.LookupParameter("Дата утверждения листа");
                            if (dateParam == null)
                            {
                                continue;
                            }
                            dateParam.Set(dateTextBox.Text);                            
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
