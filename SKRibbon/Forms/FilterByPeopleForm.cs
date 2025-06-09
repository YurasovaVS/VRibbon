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
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using Autodesk.Revit.DB.Architecture;
using System.Xml.Linq;
using Autodesk.Revit.Attributes;
using static SKRibbon.FormDesign;

namespace FilterByPeople
{
    [Transaction(TransactionMode.Manual)]
    public partial class FilterByPeopleForm : VForm
    {
        Document Doc;
        UIDocument UiDoc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        WinForms.ComboBox namesCB = new WinForms.ComboBox();
        WinForms.ComboBox paramCB = new WinForms.ComboBox();
        CheckBox checkBox = new CheckBox();

        Dictionary<string, HashSet<ElementId>> Creators =  new Dictionary <string, HashSet<ElementId>>();
        Dictionary<string, HashSet<ElementId>> LastChangedBy = new Dictionary<string, HashSet<ElementId>>();
        Dictionary<string, HashSet<ElementId>> Owners = new Dictionary<string, HashSet<ElementId>>();

        public FilterByPeopleForm(UIDocument uiDoc)
        {
            InitializeComponent();
            Doc = uiDoc.Document;
            UiDoc = uiDoc;

            // Инициализация formWrapper'а
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;

            // Обработка выделения
            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedElementIds = selection.GetElementIds();

            foreach (ElementId elementId in selectedElementIds)
            {
                Element element = Doc.GetElement(elementId);
                WorksharingTooltipInfo info = WorksharingUtils.GetWorksharingTooltipInfo(Doc, elementId);

                // Добавляем элемент в словарь создателей
                if ((info.Creator != null) && (info.Creator != ""))
                {
                    if (!Creators.ContainsKey(info.Creator))
                    {
                        HashSet<ElementId> elementIds = new HashSet<ElementId>();
                        Creators.Add(info.Creator, elementIds);
                    }
                    Creators[info.Creator].Add(element.Id);
                }
                // Добавляем элемент в словарь изменятелей
                if ((info.LastChangedBy != null) && (info.LastChangedBy!= ""))
                {
                    if (!LastChangedBy.ContainsKey(info.LastChangedBy))
                    {
                        HashSet<ElementId> elementIds = new HashSet<ElementId>();
                        LastChangedBy.Add(info.LastChangedBy, elementIds);
                    }
                    LastChangedBy[info.LastChangedBy].Add(element.Id);
                }
                // Добавляем элемент в словарь заемщиков
                if ((info.Owner!= null) && (info.Owner!= ""))
                {
                    if (!Owners.ContainsKey(info.Owner))
                    {
                        HashSet<ElementId> elementIds = new HashSet<ElementId>();
                        Owners.Add(info.Owner, elementIds);
                    }
                    Owners[info.Owner].Add(element.Id);
                }
            }

            // Инициализация выпадающего списка параметра (Создатель/Изменил/Заемщик)
            paramCB.Anchor = AnchorStyles.Left;
            paramCB.Size = new Size(200, 30);
            if (Creators.Count != 0) paramCB.Items.Add("Создал");
            if (LastChangedBy.Count != 0) paramCB.Items.Add("Изменил");
            if (Owners.Count != 0) paramCB.Items.Add("Заемщик");
            paramCB.DropDownClosed += OnParameterChanged;
            paramCB.SelectedIndex = 0;

            // Инициализация выпадающего списка имен
            namesCB.Anchor = AnchorStyles.Left;
            namesCB.Size = new Size(200, 30);
            ChooseDictionary(paramCB.SelectedItem.ToString());

            // Добавление галочки (изолировать выделение или нет)
            checkBox.Anchor = AnchorStyles.Left;
            checkBox.Text = "Изолировать выделение";
            checkBox.Checked = true;
            checkBox.Size = new Size(200, 30);

            // Добавление кнопки
            WinForms.Button okButton = new WinForms.Button();
            okButton.Anchor = AnchorStyles.Top;
            okButton.Size = new System.Drawing.Size(200, 50);
            okButton.Text = "Выделить";
            okButton.Click += RunFilter;

            // Добавление элементов во formWrapper
            paramCB.Parent = formWrapper;
            namesCB.Parent = formWrapper;
            checkBox.Parent = formWrapper;
            okButton.Parent = formWrapper;

            formWrapper.Controls.Add(paramCB);
            formWrapper.Controls.Add(namesCB);
            formWrapper.Controls.Add(checkBox);
            formWrapper.Controls.Add(okButton);

            // Добавление formWrapper'a в форму
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            this.Size = new System.Drawing.Size(250, 200);
        }
        // 
        public void RepopulateNames(string[] keys)
        {
            namesCB.Items.Clear();
            foreach (string key in keys)
            {
                namesCB.Items.Add(key);
            }
            namesCB.SelectedIndex = 0;
        }

        public void ChooseDictionary(string dictionary)
        {
            switch (dictionary)
            {
                case "Создал":
                    RepopulateNames(Creators.Keys.ToArray());
                    break;

                case "Изменил":
                    RepopulateNames(LastChangedBy.Keys.ToArray());
                    break;

                case "Заемщик":
                    RepopulateNames(Owners.Keys.ToArray());
                    break;
            }
        }
        public void OnParameterChanged(object sender, EventArgs e)
        {
            WinForms.ComboBox paramCB = (WinForms.ComboBox)sender;
            string dictionary = paramCB.SelectedItem as string;
            ChooseDictionary(dictionary);
        }

        public void RunFilter (object sender, EventArgs e)
        {
            switch (paramCB.SelectedItem.ToString())
            {
                case "Создал":
                    UiDoc.Selection.SetElementIds(Creators[namesCB.SelectedItem.ToString()]);
                    break;
                case "Изменил":
                    UiDoc.Selection.SetElementIds(LastChangedBy[namesCB.SelectedItem.ToString()]);
                    break;
                case "Заемщик":
                    UiDoc.Selection.SetElementIds(Owners[namesCB.SelectedItem.ToString()]);
                    break;
            }
            if (checkBox.Checked)
            {
                Transaction t = new Transaction(Doc, "Изолировать выделение");
                t.Start();
                Doc.ActiveView.IsolateElementsTemporary(UiDoc.Selection.GetElementIds());
                t.Commit();
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
