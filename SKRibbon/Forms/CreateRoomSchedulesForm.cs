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
using WinForm = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    public partial class CreateRoomSchedulesForm : WinForm.Form 
    {
        Document Doc;
        public CreateRoomSchedulesForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            this.AutoSize = true;
            this.AutoScroll = true;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            //Создаем Wrapper для содержимого формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.BorderStyle = BorderStyle.FixedSingle;
            formWrapper.Padding = new Padding(5, 5, 5, 5);

            //Создаем первый заголовок
            Label header1 = new Label();
            header1.Parent = formWrapper;
            formWrapper.Controls.Add(header1);
            header1.Anchor = AnchorStyles.Top;
            header1.Size = new Size(500, 30);
            header1.Text = "Выберите уровни:";

            //Создаем список уровней

            CheckedListBox levelsListBox = new CheckedListBox();
            levelsListBox.Parent = formWrapper;
            formWrapper.Controls.Add(levelsListBox);
            levelsListBox.Anchor = AnchorStyles.Top;
            levelsListBox.Size = new Size(400, 400);

            ICollection<Element> levels = new FilteredElementCollector(doc).
                                                OfCategory(BuiltInCategory.OST_Levels).
                                                WhereElementIsNotElementType().
                                                ToElements();
            foreach (Element level in levels) { 
                levelsListBox.Items.Add(level.Name);
            }

            // Добавить кнопку "Инвертировать выделение"
            Button invertButton = new Button();
            invertButton.Parent = formWrapper;
            formWrapper.Controls.Add(invertButton);
            invertButton.Anchor = AnchorStyles.Top;
            invertButton.Size = new Size(150, 30);
            invertButton.Text = "Инвертировать выделение";

            // Добавить кнопку "Создать спецификации"


        }
    }
}
