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
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static SKRibbon.FormDesign;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace SKRibbon
{
    public partial class EditParameterForm : VForm
    {
        Document Doc;

        //VTextBox ParameterName = new VTextBox();
        VComboBox ParameterName = new VComboBox();
        VTextBox PrefixText = new VTextBox();
        VTextBox SuffixText = new VTextBox();
        VTextBox FindText = new VTextBox();
        VTextBox ReplaceText = new VTextBox();
        CheckBox IsTypeParam = new CheckBox();
        

        ICollection<ElementId> SelectionIds;

        VButton OkButton = new VButton();

        int LWidth = 150;
        int RWidth = 400;
        int RowHeight = 25;

        int BorderMargin = 20;
        int ColumnSpace = 20;
        int RowSpace = 20;


        public EditParameterForm(Document doc, ICollection<ElementId> selectionIds)
        {

            Doc = doc;
            SelectionIds = selectionIds;
            InitializeComponent();
            this.Width = RWidth + LWidth + 2 * BorderMargin + ColumnSpace;

            PrefixText.Multiline = false;
            SuffixText.Multiline = false;
            FindText.Multiline = false;
            ReplaceText.Multiline = false;

            IsTypeParam.Text = "Это параметр типа";
            IsTypeParam.Size = new Size(LWidth+RWidth, RowHeight);
            IsTypeParam.Margin = new Padding(BorderMargin, 0, 0, RowSpace);
            IsTypeParam.CheckedChanged += IsTypeParam_CheckedChanged;




            // Обертка для всей формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Заголовок
            VHeaderLabel header = new VHeaderLabel();
            header.Size = new Size(RWidth + LWidth, RowHeight * 2);
            header.Text = "Отредактировать текст в параметре";
            header.Anchor = AnchorStyles.Top;


            // Лейблы для строк
            Label paramNameLabel = CreateLineLabel("Имя параметра");
            Label prefixLabel = CreateLineLabel("Добавить в начале");
            Label suffixNameLabel = CreateLineLabel("Добавить в конце");
            Label findNameLabel = CreateLineLabel("Найти");
            Label replaceNameLabel = CreateLineLabel("Заменить на");

            // Поля для строк
            ParameterName.Size = new Size(RWidth, RowHeight);
            PrefixText.Size = new Size(RWidth, RowHeight);
            SuffixText.Size = new Size(RWidth, RowHeight);
            FindText.Size = new Size(RWidth, RowHeight);
            ReplaceText.Size = new Size(RWidth, RowHeight);


            FlowLayoutPanel paramNameBorder = new FlowLayoutPanel();
            paramNameBorder.AutoSize = true;
            paramNameBorder.BorderStyle = BorderStyle.FixedSingle;

            paramNameBorder.Controls.Add(ParameterName);
            ParameterName.Parent = paramNameBorder;

            //ParameterName.Text = "Наименование";
            FillParameterList(IsTypeParam.Checked);
            PrefixText.Text = "";
            SuffixText.Text = "";
            FindText.Text = "";
            ReplaceText.Text = "";

            //Обертки для строк
            FlowLayoutPanel paramNamePanel = new FlowLayoutPanel();
            paramNamePanel.AutoSize = true;
            paramNamePanel.FlowDirection = FlowDirection.LeftToRight;


            FlowLayoutPanel prefixPanel = new FlowLayoutPanel();
            prefixPanel.AutoSize = true;
            prefixPanel.FlowDirection = FlowDirection.LeftToRight;

            FlowLayoutPanel suffixPanel = new FlowLayoutPanel();
            suffixPanel.AutoSize = true;
            suffixPanel.FlowDirection = FlowDirection.LeftToRight;

            FlowLayoutPanel findPanel = new FlowLayoutPanel();
            findPanel.AutoSize = true;
            findPanel.FlowDirection = FlowDirection.LeftToRight;

            FlowLayoutPanel replacePanel = new FlowLayoutPanel();
            replacePanel.AutoSize = true;
            replacePanel.FlowDirection = FlowDirection.LeftToRight;

            // Собираем строки
            paramNameLabel.Parent = paramNamePanel;
            paramNamePanel.Controls.Add(paramNameLabel);
            paramNameBorder.Parent = paramNamePanel;
            paramNamePanel.Controls.Add(paramNameBorder);

            prefixLabel.Parent = prefixPanel;
            prefixPanel.Controls.Add(prefixLabel);
            PrefixText.Parent = prefixPanel;
            prefixPanel.Controls.Add(PrefixText);

            suffixNameLabel.Parent = suffixPanel;
            suffixPanel.Controls.Add(suffixNameLabel);
            SuffixText.Parent = suffixPanel;
            suffixPanel.Controls.Add(SuffixText);

            findNameLabel.Parent = findPanel;
            findPanel.Controls.Add(findNameLabel);
            FindText.Parent = findPanel;
            findPanel.Controls.Add(FindText);

            replaceNameLabel.Parent = replacePanel;
            replacePanel.Controls.Add(replaceNameLabel);
            ReplaceText.Parent = replacePanel;
            replacePanel.Controls.Add(ReplaceText);

            // Кнопка
            OkButton.Size = new Size(RWidth + LWidth + 2* BorderMargin, RowHeight * 2);
            OkButton.Text = "Запустить";
            OkButton.Click += RunChanges;

            // Добавляем строки в форму

            header.Parent = formWrapper;
            formWrapper.Controls.Add(header);

            paramNamePanel.Parent = formWrapper;
            formWrapper.Controls.Add(paramNamePanel);

            IsTypeParam.Parent = formWrapper;
            formWrapper.Controls.Add(IsTypeParam);

            prefixPanel.Parent = formWrapper;
            formWrapper.Controls.Add(prefixPanel);

            suffixPanel.Parent = formWrapper;
            formWrapper.Controls.Add(suffixPanel);

            findPanel.Parent = formWrapper;
            formWrapper.Controls.Add(findPanel);

            replacePanel.Parent = formWrapper;
            formWrapper.Controls.Add(replacePanel);

            OkButton.Parent = formWrapper;
            formWrapper.Controls.Add(OkButton);

            this.Height = formWrapper.Height + 10;
        }

        private void IsTypeParam_CheckedChanged(object sender, EventArgs e)
        {
            FillParameterList(IsTypeParam.Checked);
        }


        private void RunChanges(object sender, EventArgs e)
        {
            if (ParameterName.Text == "")
            {
                TaskDialog.Show("Параметр", "Введите имя параметра");
            }
            else
            {
                Transaction t = new Transaction(Doc, "Поиск и замена");
                t.Start();

                int goodCounter = 0;
                int badCounter = 0;

                if (!IsTypeParam.Checked)
                {
                    foreach (ElementId elementId in SelectionIds)
                    {
                        Element el = Doc.GetElement(elementId);
                        Parameter param = el.LookupParameter(ParameterName.Text);

                        if (param != null)
                        {
                            string paramText = param.AsValueString();
                            if (FindText.Text.Length > 0)
                            {
                                paramText = paramText.Replace(FindText.Text, ReplaceText.Text);
                            }
                            param.Set(PrefixText.Text + paramText + SuffixText.Text);
                            goodCounter++;

                        }
                        else
                        {
                            badCounter++;
                        }
                    }
                }
                else
                {
                    HashSet<ElementId> elementTypes = new HashSet<ElementId>();
                    foreach (ElementId elementId in SelectionIds)
                    {
                        Element el = Doc.GetElement(elementId);
                        ElementId elTypeId = el.GetTypeId();
                        elementTypes.Add(elTypeId);
                    }
                    foreach (ElementId elTypeId in elementTypes)
                    {
                        Element elType = Doc.GetElement(elTypeId);
                        Parameter param = elType.LookupParameter(ParameterName.Text);

                        if (param != null)
                        {
                            string paramText = param.AsValueString();
                            if (FindText.Text.Length > 0)
                            {
                                paramText = paramText.Replace(FindText.Text, ReplaceText.Text);
                            }
                            param.Set(PrefixText.Text + paramText + SuffixText.Text);
                            goodCounter++;

                        }
                        else
                        {
                            badCounter++;
                        }
                    }
                }                    

                
                t.Commit();
                TaskDialog.Show("Задача выполнена", "Успешно заменено: " + goodCounter.ToString() + " элементов; ошибок: " + badCounter.ToString());
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public Label CreateLineLabel(string name)
        {
            Label label = new Label();
            label.Text = name;
            label.Size = new Size(LWidth, RowHeight);
            label.Margin = new Padding(BorderMargin, 0, 0, 0);
            label.Font = new Font(Label.DefaultFont, FontStyle.Bold);

            return label;
        }

        public void FillParameterList (bool isTypeParam)
        {
            HashSet<string> paramNames = new HashSet<string>();
            HashSet<ElementId> elementIds = new HashSet<ElementId>();

            if (isTypeParam)
            {
                HashSet<ElementId> elementTypes = new HashSet<ElementId>();
                foreach (ElementId elementId in SelectionIds)
                {
                    Element el = Doc.GetElement(elementId);
                    ElementId elTypeId = el.GetTypeId();
                    elementIds.Add(elTypeId);
                }
            }
            else
            {
                foreach (ElementId elId in SelectionIds)
                {
                    elementIds.Add(elId);
                }                
            }

            bool isFirst = true;
            foreach (ElementId elementId in elementIds)
            {
                Element element = Doc.GetElement(elementId);
                if (isFirst)
                {
                    foreach (Parameter param in element.Parameters)
                    {
                        if ((!param.IsReadOnly) &&
                            ((param.Definition.GetDataType() == SpecTypeId.String.Text) ||
                            (param.Definition.GetDataType() == SpecTypeId.String.MultilineText))) paramNames.Add(param.Definition.Name);
                    }
                    isFirst = false;
                }
                else
                {
                    HashSet<string> temp = new HashSet<string>();
                    foreach (Parameter param in element.Parameters)
                    {
                        if ((!param.IsReadOnly) && 
                            ((param.Definition.GetDataType() == SpecTypeId.String.Text) || 
                            (param.Definition.GetDataType() == SpecTypeId.String.MultilineText))) temp.Add(param.Definition.Name);
                    }
                    paramNames.IntersectWith(temp);
                }
            }

            ParameterName.Items.Clear();
            if (paramNames.Count > 0)
            {
                ParameterName.Enabled = true;
                foreach (string name in paramNames)
                {
                    ParameterName.Items.Add(name);
                }
            }
            else
            {
                ParameterName.Enabled = false;
                ParameterName.Items.Add("Доступные параметры не найдены");
            }
            ParameterName.SelectedIndex = 0;
        }
    }    
}
