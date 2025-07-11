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
using System.Collections;
using System.Net.NetworkInformation;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    
    public partial class FixWorkGroupsForm : VForm
    {
        Dictionary<BuiltInCategory, string> CategoryNames = new Dictionary<BuiltInCategory, string>()
        {
            { BuiltInCategory.OST_CurtaSystem, "Витражи"},
            { BuiltInCategory.OST_Doors, "Двери"},
            { BuiltInCategory.OST_Roofs, "Кровля"},
            { BuiltInCategory.OST_Stairs, "Лестницы"},
            { BuiltInCategory.OST_PlumbingFixtures, "Оборудование сантех."},
            { BuiltInCategory.OST_DuctTerminal, "Оборудование сантех."},
            { BuiltInCategory.OST_MechanicalEquipment, "Оборудование мех."},
            { BuiltInCategory.OST_Railings, "Ограждения"},
            { BuiltInCategory.OST_StairsRailing, "Ограждения"},
            { BuiltInCategory.OST_Windows, "Окна"},
            { BuiltInCategory.OST_ShaftOpening, "Отверстия"},
            { BuiltInCategory.OST_Grids, "Оси"},
            { BuiltInCategory.OST_Floors, "Полы"},
            { BuiltInCategory.OST_Rooms, "Помещения"},
            { BuiltInCategory.OST_RoomSeparationLines, "Помещения"},
            { BuiltInCategory.OST_Ceilings, "Потолки"},
            { BuiltInCategory.OST_Levels, "Уровни"},
            { BuiltInCategory.OST_Mass, "Формы"}
        };

        //Dictionary<string, BuiltInCategory> NamesCategory = new Dictionary<string, BuiltInCategory>();
        HashSet<string> NamesCategory = new HashSet<string>();

        Dictionary<string, Workset> Name_Workset = new Dictionary<string, Workset>();
        HashSet<string> Errors = new HashSet<string>();

        Document Doc;

        WinForms.FlowLayoutPanel formWrapper = new WinForms.FlowLayoutPanel();
        WinForms.FlowLayoutPanel CategoriesWrapper = new WinForms.FlowLayoutPanel();
        WinForms.FlowLayoutPanel CategoriesByTypeWrapper = new WinForms.FlowLayoutPanel();

        int LabelWidth = 150;
        int LabelHeight = 20;

        int TextBoxWidth = 150;
        int TextBoxHeight = 20;

        int ComboBoxWidth = 200;
        int ComboBoxHeight = 20;

        int MarginLeft = 35;

        public FixWorkGroupsForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            this.AutoScroll = true;
            this.Width = LabelWidth + TextBoxWidth + ComboBoxWidth + MarginLeft;
            this.Height = 480;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            VHeaderLabel byTypeNameLabel = new VHeaderLabel();
            byTypeNameLabel.Size = new Size(LabelWidth + TextBoxWidth + ComboBoxWidth + MarginLeft, 35);
            byTypeNameLabel.Text = "Распределение по имени типоразмера";

            byTypeNameLabel.Parent = formWrapper;
            formWrapper.Controls.Add(byTypeNameLabel);

            
            CategoriesByTypeWrapper.AutoSize = true;
            CategoriesByTypeWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            CategoriesByTypeWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(CategoriesByTypeWrapper);
            
            FlowLayoutPanel ConstrFLP = WorkgroupAndFamilyFLP("КЖ", "КЖ");
            FlowLayoutPanel DisabledPeopleFLP = WorkgroupAndFamilyFLP("ОДИ", "ОДИ");
            DisabledPeopleFLP.BackColor = System.Drawing.Color.FromArgb(255, 243, 221, 255);
            FlowLayoutPanel WallsInternalFLP = WorkgroupAndFamilyFLP("Внутренние стены", "АР_В");
            FlowLayoutPanel WallsExternalFLP = WorkgroupAndFamilyFLP("Наружные стены", "АР_Н");
            WallsExternalFLP.BackColor = System.Drawing.Color.FromArgb(255, 243, 221, 255);
            FlowLayoutPanel WallsDesignFLP = WorkgroupAndFamilyFLP("Отделка", "АР_О");
            FlowLayoutPanel CurtainWalls = WorkgroupAndFamilyFLP("Витражи", "витраж");
            CurtainWalls.BackColor = System.Drawing.Color.FromArgb(255, 243, 221, 255);
            FlowLayoutPanel GenPlan = WorkgroupAndFamilyFLP("Генплан", "ГП");



            IEnumerable<Workset> worksetsIE = new FilteredWorksetCollector(Doc).OfKind(WorksetKind.UserWorkset).ToWorksets();

            int i = 0;
            foreach (Workset workset in worksetsIE) {
                Name_Workset.Add(workset.Name, workset);

                // Заполняем существующие выпадающие списки
                foreach (var panel in CategoriesByTypeWrapper.Controls)
                {
                    FlowLayoutPanel p = panel as FlowLayoutPanel;
                    WinForms.FlowLayoutPanel comboPanel = p.Controls[2] as WinForms.FlowLayoutPanel;
                    VComboBox box = comboPanel.Controls[0] as VComboBox;
                    WinForms.Label label = p.Controls[0] as WinForms.Label;
                    box.Items.Add(workset.Name);
                    var words = label.Text.Split(' ');
                    if (workset.Name.Contains(words[0])) box.SelectedIndex = i;
                }
                i++;
            }

            // Добавляем блок с категориями
            // Обертка
            CategoriesWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            CategoriesWrapper.AutoSize = true;

            // Заголовок
            VHeaderLabel byTypeLabel = new VHeaderLabel();
            byTypeLabel.Size = new Size(LabelWidth + TextBoxWidth + ComboBoxWidth + MarginLeft, 35);
            byTypeLabel.Text = "Распределение по типу элемента";

            byTypeLabel.Parent = formWrapper;
            formWrapper.Controls.Add(byTypeLabel);

            // Заполнение из словаря
            int counter = 0;
            foreach (KeyValuePair<BuiltInCategory, string> category in CategoryNames) {
                if (NamesCategory.Add(category.Value))
                {
                    WinForms.FlowLayoutPanel catPanel = WorkgroupFLP(category.Value);
                    if (counter % 2 != 0) catPanel.BackColor = System.Drawing.Color.FromArgb(255, 243, 221, 255);
                    counter++;

                    catPanel.Parent = CategoriesWrapper;
                    CategoriesWrapper.Controls.Add(catPanel);
                }                
            }

            formWrapper.Controls.Add(CategoriesWrapper);
            CategoriesWrapper.Parent = formWrapper;

            // Кнопка
            VButton OkButton = new VButton();
            OkButton.Size = new Size(ComboBoxWidth, 50);
            OkButton.Text = "ОК";
            OkButton.Anchor = AnchorStyles.Top;
            OkButton.Margin = new Padding(0, 10, 0, 0);
            OkButton.Click += RunFixing;

            formWrapper.Controls.Add(OkButton);
            OkButton.Parent = formWrapper;
        }

        private void RunFixing(object sender, EventArgs e)
        {
            ICollection<Element> elements = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements();
            StringBuilder sb_ReadOnly = new StringBuilder();
            StringBuilder sb_TypeNotSupported = new StringBuilder();
            StringBuilder sb_ElementHasNoCategory = new StringBuilder();
            sb_ReadOnly.AppendLine("Рабочий набор данных элементов стоит в режиме \"Только для чтения\": ");
            sb_TypeNotSupported.AppendLine("Тип данных элементов не поддерживается. Обратитесь к разработчику:");
            sb_ElementHasNoCategory.AppendLine("У данных элементов отсутствует категория: ");

            Transaction t = new Transaction(Doc, "Исправить рабочие наборы");
            t.Start();

            foreach (Element element in elements) {

                if (element.Id.ToString() == "1212844")
                {
                    bool flag = true;
                }

                // Проверяем, существует ли такой параметр, и не ReadOnly ли он
                Parameter worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                if (worksetParam == null) continue;
                if (worksetParam.IsReadOnly) {
                    sb_ReadOnly.AppendLine(element.Name + " " + element.Id);
                    continue;
                }
                

                string worksetName = "";

                bool skipFurtherChecks = false;

                // Проверяем, не относится ли он к рабочим наборам по типу (КЖ, ОДИ, внутренние и внешние стены, геплан)
                foreach (var catByTypeWrapper in CategoriesByTypeWrapper.Controls)
                {                    
                    if (skipFurtherChecks) continue;

                    FlowLayoutPanel panel = catByTypeWrapper as FlowLayoutPanel;
                    WinForms.TextBox text = panel.Controls[1] as WinForms.TextBox;
                    WinForms.FlowLayoutPanel comboPanel = panel.Controls[2] as WinForms.FlowLayoutPanel;
                   VComboBox combo = comboPanel.Controls[0] as VComboBox;                                     

                    if (element.Name.Contains(text.Text)) {
                        worksetName = combo.Text;
                        skipFurtherChecks = true;
                    }
                }
                if (!skipFurtherChecks)
                {
                    // Проверяем, есть ли у данного элемента категория
                    if (element.Category == null)
                    {
                        sb_ElementHasNoCategory.AppendLine(element.Name + " " + element.Id);
                        continue;
                    }
                    // Проверяем, есть ли такая категория в нашем словаре
#if DEBUG2021 || REVIT2021
                    if (!CategoryNames.ContainsKey((BuiltInCategory)element.Category.Id.IntegerValue))
                    {
                        sb_TypeNotSupported.AppendLine(element.Name + " " + element.Id + (BuiltInCategory)element.Category.Id.IntegerValue);
                        continue;
                    }
#elif DEBUG2024 || REVIT2024
                    if (!CategoryNames.ContainsKey((BuiltInCategory)element.Category.Id.Value))
                    {
                        sb_TypeNotSupported.AppendLine(element.Name + " " + element.Id + (BuiltInCategory)element.Category.Id.Value);
                        continue;
                    }
#endif

                    // Смотрим, к какому рабочему набору относится элемент (BuiltInCategory)
                    foreach (var categoryFLP in CategoriesWrapper.Controls)
                    {
                        if (skipFurtherChecks) continue; // Если мы уже перенесли элемент в какой-либо рабочий набор,              


                        FlowLayoutPanel panel = categoryFLP as FlowLayoutPanel;
                        WinForms.Label catName = panel.Controls[0] as WinForms.Label;
                        WinForms.FlowLayoutPanel comboPanel = panel.Controls[1] as WinForms.FlowLayoutPanel;
                        VComboBox combo = comboPanel.Controls[0] as VComboBox;

#if DEBUG2021 || REVIT2021
                        if (CategoryNames[(BuiltInCategory)element.Category.Id.IntegerValue] != catName.Text) continue;
#elif DEBUG2024 || REVIT2024
                        if (CategoryNames[(BuiltInCategory)element.Category.Id.Value] != catName.Text) continue;
#endif                        
                        worksetName = combo.Text;
                        skipFurtherChecks = true;
                    }
                }                
                if (worksetName == "") continue; // Если по какой-то причине worksetName остается не задан, пропускаем элемент
                worksetParam.Set(Name_Workset[worksetName].Id.IntegerValue);
            } // Конец цикла foreach (Element element in elements)

            t.Commit();

            TaskDialog.Show("Ошибки", sb_ReadOnly.ToString());
            TaskDialog.Show("Ошибки", sb_TypeNotSupported.ToString());
            // TaskDialog.Show("Ошибки", sb_ElementHasNoCategory.ToString());           
                

            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
        }

        // Создание группы для проверки по категории
        public WinForms.FlowLayoutPanel WorkgroupFLP(string Name)
        {
            // [0] Наименование элемента
            // [1] Выпадающий список с рабочими наборами


            // [0] Наименование элемента
            WinForms.Label workgroupName = new WinForms.Label();
            workgroupName.Text = Name;
            workgroupName.Size = new Size(LabelWidth, LabelHeight);
            workgroupName.Font = new Font(Label.DefaultFont, FontStyle.Bold);
            workgroupName.TextAlign = ContentAlignment.MiddleLeft;
            workgroupName.Margin = new Padding(MarginLeft, 0, 0, 0);

            // [1] Выпадающий список с рабочими наборами
            // Обертка
            WinForms.FlowLayoutPanel comboBoxWrapper = new WinForms.FlowLayoutPanel();
            comboBoxWrapper.AutoSize = true;
            comboBoxWrapper.BorderStyle = BorderStyle.FixedSingle;

            // Список
            VComboBox workgroupsCB = new VComboBox();
            workgroupsCB.Size = new Size(ComboBoxWidth + TextBoxWidth, ComboBoxHeight);
            foreach (KeyValuePair<string, Workset> workset in Name_Workset)
            {
                int i = workgroupsCB.Items.Add(workset.Key);
                var words = Name.Split(' ');
                if (workset.Key.Contains(words[0])) workgroupsCB.SelectedIndex = i;
            }

            workgroupsCB.Parent = comboBoxWrapper;
            comboBoxWrapper.Controls.Add(workgroupsCB);

            // Общая обертка
            WinForms.FlowLayoutPanel workgroupFLP = new WinForms.FlowLayoutPanel();
            workgroupFLP.AutoSize = true;
            workgroupFLP.FlowDirection = WinForms.FlowDirection.LeftToRight;
            workgroupFLP.Margin = new Padding(10, 0, 0, 0);

            workgroupFLP.Controls.Add(workgroupName);
            workgroupFLP.Controls.Add(comboBoxWrapper);
            workgroupName.Parent = workgroupFLP;
            comboBoxWrapper.Parent = workgroupFLP;
                        
            return workgroupFLP;
        }

        // Создание группы для проверки по семейству
        public WinForms.FlowLayoutPanel WorkgroupAndFamilyFLP(string Name, string TextBoxText)
        {
            // [0] Наименование элемента
            // [1] Поле для сравнения с типом семейства
            // [2] Выпадающий список с рабочими наборами

            // [0] Наименование элемента
            WinForms.Label ConstrLabel = new WinForms.Label();
            ConstrLabel.Size = new Size(LabelWidth, LabelHeight);
            ConstrLabel.Text = Name;
            ConstrLabel.Font = new Font(Label.DefaultFont, FontStyle.Bold);
            ConstrLabel.TextAlign = ContentAlignment.MiddleLeft;
            ConstrLabel.Margin = new Padding(MarginLeft, 0, 0, 0);

            // [1] Поле для сравнения с типом семейства
            WinForms.TextBox FLPTextBox = new WinForms.TextBox();
            FLPTextBox.Size = new Size(TextBoxWidth, TextBoxHeight);
            FLPTextBox.Text = TextBoxText;

            // [2] Выпадающий список с рабочими наборами
            // Обертка
            WinForms.FlowLayoutPanel comboBoxWrapper = new WinForms.FlowLayoutPanel();
            comboBoxWrapper.AutoSize = true;
            comboBoxWrapper.BorderStyle = BorderStyle.FixedSingle;

            // Выпадающий список
            VComboBox FLPComboBox = new VComboBox();
            FLPComboBox.Size = new Size(ComboBoxWidth, ComboBoxHeight);

            FLPComboBox.Parent = comboBoxWrapper;
            comboBoxWrapper.Controls.Add(FLPComboBox);

            //------------
            // Общая обертка

            WinForms.FlowLayoutPanel FLPanel = new WinForms.FlowLayoutPanel();
            FLPanel.FlowDirection = WinForms.FlowDirection.LeftToRight;
            FLPanel.AutoSize = true;
            FLPanel.Margin = new Padding(10, 0, 0, 0);
            
            FLPanel.Controls.Add(ConstrLabel);
            FLPanel.Controls.Add(FLPTextBox);
            FLPanel.Controls.Add(comboBoxWrapper);

            ConstrLabel.Parent = FLPanel;
            FLPTextBox.Parent = FLPanel;
            comboBoxWrapper.Parent = FLPanel;

            CategoriesByTypeWrapper.Controls.Add(FLPanel);
            FLPanel.Parent = CategoriesByTypeWrapper;

            return FLPanel;
        }
    }   
}
