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
using Autodesk.Revit.Attributes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Excel;

namespace CopyListsTree
{
    public partial class CopyListsForm : System.Windows.Forms.Form
    {
        Document Doc;
        HashSet<string> sheetNumbers = new HashSet<string>();
        HashSet<string> viewNames = new HashSet<string>();
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        public CopyListsForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            this.AutoScroll = true;
            this.Width = 1150;
            this.Height = 900;

            //Собираем все вьюпорты в проекте и запихиваем их имена во viewNames
            ICollection<Element> views = new FilteredElementCollector(Doc).
                                                OfClass(typeof(Autodesk.Revit.DB.View)).
                                                WhereElementIsNotElementType().
                                                ToElements();
            foreach (Autodesk.Revit.DB.View view in views)
            {
                viewNames.Add(view.Name);
            }

            //Собираем все листы в основу для древа
            /* Словарь Зданий
             *      Здание : Словарь Томов
             *          Том : Список объектов
             *              Объект ViewSheet
             *                  
            */

            ICollection<Element> sheets = new FilteredElementCollector(Doc).
                                            OfCategory(BuiltInCategory.OST_Sheets).
                                            WhereElementIsNotElementType().
                                            ToElements();
            foreach (ViewSheet sheet in sheets)
            {
                sheetNumbers.Add(sheet.SheetNumber); // Создаем множество номеров листов

                Parameter tomeParam = sheet.LookupParameter("ADSK_Штамп Раздел проекта");
                Parameter buildingParam = sheet.LookupParameter("ADSK_Примечание");

                if (tomeParam == null) tomeParam = sheet.LookupParameter("Раздел проекта");
                if (buildingParam == null) buildingParam = sheet.LookupParameter("Примечание");

                if (tomeParam != null && buildingParam != null)
                {
                    string tome = "<РАЗДЕЛ НЕ ЗАДАН>";
                    string building = "<ПРИМЕЧАНИЕ НЕ ЗАДАНО>";
                    if (tomeParam.AsString() != null && tomeParam.AsString() != "") 
                    { 
                        tome = tomeParam.AsString();
                    }
                    if (buildingParam.AsString() != null && buildingParam.AsString() != "")
                    {
                        building = buildingParam.AsString();
                    }
                    // Если в первом словаре нет такого здания, создаем его
                    if (!buildingsDict.ContainsKey(building))
                    {
                        Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                        buildingsDict.Add(building, tomesDict);
                    }
                    // Если во вложенном словаре здания нет такого тома, создаем его
                    if (!buildingsDict[building].ContainsKey(tome))
                    {
                        List<ViewSheet>sheetsList = new List<ViewSheet>();
                        buildingsDict[building].Add(tome, sheetsList);
                    }
                    // Добавляем лист в нужный том
                    buildingsDict[building][tome].Add(sheet);
                }
            } // Конец создания словаря листов

            //Создаем wrapper для всей формы
            WinForms.FlowLayoutPanel formWrapper = new WinForms.FlowLayoutPanel();
            formWrapper.BorderStyle = WinForms.BorderStyle.FixedSingle;
            formWrapper.FlowDirection = WinForms.FlowDirection.LeftToRight;
            formWrapper.AutoSize = true;
            this.Controls.Add(formWrapper);
            formWrapper.Parent = this;

            //Создаем wrapper для дерева
            /*  Структура:
             *  treeWrapper
             *      buildingWrapper
             *          buildingHeader
             *          tomesWrapper
             *              tomeWrapper
             *                  tomeHeader
             *                  sheetsWrapper
             *                      sheetWrapper
             *                          button
             *                          templatesWrapper
             *                              templateWrapper
             *                                  button + textbox + textbox + textbox + textbox
             *   По-русски:
             *   Обертка дерева (зданий)
             *      Обертка здания
             *          Заголовок здания
             *          Обертка томов
             *              Обертка тома
             *                  Заголовок тома
             *                  Обертка листов
             *                      Обертка листа
             *                          Заголовок листа (Кнопка)
             *                          Обертка шаблонов новых листов
             *                              Обертка шаблона новых листов
             *                                  Шаблон (Кнопка + Поле 1 + Поле 2 + Поле 3 + Поле 4)
             * 
             */

            WinForms.FlowLayoutPanel treeWrapper = new WinForms.FlowLayoutPanel();
            treeWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(treeWrapper);
            treeWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            treeWrapper.AutoSize = true;
            treeWrapper.Anchor = WinForms.AnchorStyles.Left;
            treeWrapper.BorderStyle = WinForms.BorderStyle.FixedSingle;

            foreach (var building in buildingsDict)
            {
                // Создаем wrapper для здания
                WinForms.FlowLayoutPanel buildingWrapper = new WinForms.FlowLayoutPanel();
                buildingWrapper.Parent = treeWrapper;
                treeWrapper.Controls.Add(buildingWrapper);
                buildingWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                buildingWrapper.AutoSize = true;
                buildingWrapper.Anchor = WinForms.AnchorStyles.Left;

                // Создаем заголовок
                WinForms.Label buildingHeader = new WinForms.Label();
                buildingHeader.Parent = buildingWrapper;
                buildingWrapper.Controls.Add(buildingHeader);
                buildingHeader.Text = building.Key;
                buildingHeader.Width = 300;

                // Создаем wrapper для томов этого здания
                WinForms.FlowLayoutPanel tomesWrapper = new WinForms.FlowLayoutPanel();
                tomesWrapper.Parent = buildingWrapper;
                buildingWrapper.Controls.Add(tomesWrapper);
                tomesWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                tomesWrapper.AutoSize = true;
                tomesWrapper.Anchor = WinForms.AnchorStyles.Left;
                tomesWrapper.BorderStyle = WinForms.BorderStyle.FixedSingle;

                Dictionary<string, List<ViewSheet>> tomes = building.Value;

                foreach (var tome in tomes)
                {
                    // Создаем wrapper ТОМА
                    WinForms.FlowLayoutPanel tomeWrapper = new WinForms.FlowLayoutPanel();
                    tomeWrapper.Parent = tomesWrapper;
                    tomesWrapper.Controls.Add(tomeWrapper);
                    tomeWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                    tomeWrapper.AutoSize = true;
                    tomeWrapper.Anchor = WinForms.AnchorStyles.Left;

                    // Создаем заголовок
                    WinForms.Label tomeHeader = new WinForms.Label();
                    tomeHeader.Parent = tomeWrapper;
                    tomeWrapper.Controls.Add(tomeHeader);
                    tomeHeader.Text = tome.Key;
                    tomeHeader.Width = 300;

                    // Создаем wrapper для уже существующих листов
                    // Создаем wrapper для листов этого здания
                    WinForms.FlowLayoutPanel sheetsWrapper = new WinForms.FlowLayoutPanel();
                    sheetsWrapper.Parent = tomeWrapper;
                    tomeWrapper.Controls.Add(sheetsWrapper);
                    sheetsWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                    sheetsWrapper.AutoSize = true;
                    sheetsWrapper.Anchor = WinForms.AnchorStyles.Left;

                    List<ViewSheet> viewSheets = tome.Value;

                    viewSheets = viewSheets.OrderBy(view => view.SheetNumber).ToList<ViewSheet>();

                    foreach (ViewSheet viewSheet in viewSheets)
                    {
                        // Создаем обертку для листа
                        WinForms.FlowLayoutPanel sheetWrapper = new WinForms.FlowLayoutPanel();
                        sheetWrapper.Parent = sheetsWrapper;
                        sheetsWrapper.Controls.Add(sheetWrapper);
                        sheetWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                        sheetWrapper.AutoSize = true;
                        sheetWrapper.Anchor = WinForms.AnchorStyles.Left;

                        // Создаем кнопку с номером и названием листа
                        SheetButton button = new SheetButton();
                        button.Parent = sheetWrapper;
                        sheetWrapper.Controls.Add(button);
                        button.Width = 750;
                        button.Text = viewSheet.SheetNumber + " : " + viewSheet.Name;
                        button.Click += AddTemplate;

                        button.sheet = viewSheet;
                        button.buildingName = building.Key;
                        button.tomeName = tome.Key;

                        // Создаем wrapper для шаблонов копий листов
                        WinForms.FlowLayoutPanel templatesWrapper = new WinForms.FlowLayoutPanel();
                        templatesWrapper.Parent = sheetWrapper;
                        sheetWrapper.Controls.Add(templatesWrapper);
                        templatesWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
                        templatesWrapper.AutoSize = true;
                        templatesWrapper.Anchor = WinForms.AnchorStyles.Left;
                    }
                }
            }

            // Добавляем опции формы
            WinForms.FlowLayoutPanel optionsWrapper = new WinForms.FlowLayoutPanel();
            formWrapper.Controls.Add(optionsWrapper);
            optionsWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            optionsWrapper.AutoSize = true;
            optionsWrapper.Anchor = WinForms.AnchorStyles.Top;
            optionsWrapper.BorderStyle = WinForms.BorderStyle.FixedSingle;

            // Добавляем галочку для копирования данных штампов
            WinForms.CheckBox stampCheck = new WinForms.CheckBox();
            stampCheck.Text = "Копировать данные штампов";
            stampCheck.Width = 300;

            // Добавляем галочку для копирования видов
            WinForms.CheckBox viewsCheck = new WinForms.CheckBox();
            viewsCheck.Text = "Копировать виды";
            viewsCheck.Width = 300;

            // Добавляем галочку для копирования легенд
            WinForms.CheckBox legendsCheck = new WinForms.CheckBox();
            legendsCheck.Text = "Копировать легенды";
            legendsCheck.Width = 300;

            // Добавляем галочку для копирования спецификаций
            WinForms.CheckBox specCheck = new WinForms.CheckBox();
            specCheck.Text = "Копировать спецификации";
            specCheck.Width = 300;

            // Добавляем галочку для копирования аннотативных объектов
            WinForms.CheckBox annoCheck = new WinForms.CheckBox();
            annoCheck.Text = "Копировать аннотативные объекты";
            annoCheck.Width = 300;

            // Добавляем кнопку ОК
            WinForms.Button OKbutton = new WinForms.Button();
            OKbutton.Text = "Начать копирование";
            OKbutton.Size = new Size(300, 50);
            OKbutton.Click += CopyLists;

            // Добавляем все контролы во wrapper
            stampCheck.Parent = optionsWrapper;
            viewsCheck.Parent = optionsWrapper;
            legendsCheck.Parent = optionsWrapper;
            specCheck.Parent = optionsWrapper;
            annoCheck.Parent = optionsWrapper;
            OKbutton.Parent = optionsWrapper;

            optionsWrapper.Controls.Add(stampCheck);
            optionsWrapper.Controls.Add(viewsCheck);
            optionsWrapper.Controls.Add(legendsCheck);
            optionsWrapper.Controls.Add(specCheck);
            optionsWrapper.Controls.Add(annoCheck);
            optionsWrapper.Controls.Add(OKbutton);
        }

        //-------------------------------------------
        // 0. События
        //-------------------------------------------

        // 0.1. Добавление шаблона
        public void AddTemplate(object sender, EventArgs e)
        {
            SheetButton button = (SheetButton)sender;
            WinForms.FlowLayoutPanel sheetWrapper = (WinForms.FlowLayoutPanel)button.Parent;
            WinForms.FlowLayoutPanel templatesWrapper = (WinForms.FlowLayoutPanel)sheetWrapper.Controls[1];

            // Задаем обертку строки шаблона ( Кнопка - Поле 1 - Поле 2 - Поле 3 - Поле 4 )
            WinForms.FlowLayoutPanel templateWrapper = new WinForms.FlowLayoutPanel();            
            templateWrapper.AutoSize = true;
            templateWrapper.Anchor = WinForms.AnchorStyles.Left;
            templateWrapper.FlowDirection = WinForms.FlowDirection.LeftToRight;

            // Создаем кнопку
            WinForms.Button deleteButton = new WinForms.Button();
            deleteButton.Parent = templateWrapper;
            templateWrapper.Controls.Add(deleteButton);
            deleteButton.Text = "X";
            deleteButton.Size = new Size(20, 20);
            deleteButton.Margin = new WinForms.Padding(0, 0, 10, 0);
            deleteButton.Click += RemoveTemplate;

            // Создаем Поле 1 (ADSK_Примечание) (Номер здания)
            System.Windows.Forms.TextBox buildingTB = new System.Windows.Forms.TextBox();
            buildingTB.Parent = templateWrapper;
            templateWrapper.Controls.Add(buildingTB);
            buildingTB.Width = 100;
            buildingTB.Text = button.buildingName;
            buildingTB.Margin = new WinForms.Padding(0, 0, 10, 0);

            // Создаем Поле 2 (ADSK_Штамп Раздел проекта) (Название тома)
            System.Windows.Forms.TextBox tomeTB = new System.Windows.Forms.TextBox();
            tomeTB.Parent = templateWrapper;
            templateWrapper.Controls.Add(tomeTB);
            tomeTB.Width = 100;
            tomeTB.Text = button.tomeName;
            tomeTB.Margin = new WinForms.Padding(0, 0, 10, 0);

            // Создаем Поле З (Номер листа)
            System.Windows.Forms.TextBox numTB = new System.Windows.Forms.TextBox();
            numTB.Parent = templateWrapper;
            templateWrapper.Controls.Add(numTB);
            numTB.Width = 150;
            numTB.Text = button.sheet.SheetNumber;
            numTB.Margin = new WinForms.Padding(0, 0, 10, 0);

            // Создаем Поле 4 (Наименование листа)
            System.Windows.Forms.TextBox nameTB = new System.Windows.Forms.TextBox();
            nameTB.Parent = templateWrapper;
            templateWrapper.Controls.Add(nameTB);
            nameTB.Width = 300;
            nameTB.Text = button.sheet.Name;
            nameTB.Margin = new WinForms.Padding(0, 0, 10, 0);

            templateWrapper.Parent = templatesWrapper;
            templatesWrapper.Controls.Add(templateWrapper);
        }

        // 0.2. Удаление шаблона
        public void RemoveTemplate(object sender, EventArgs e)
        {
            WinForms.Button button = (WinForms.Button)sender;
            WinForms.FlowLayoutPanel templateWrapper = (WinForms.FlowLayoutPanel)button.Parent;
            WinForms.FlowLayoutPanel templatesWrapper = (WinForms.FlowLayoutPanel)templateWrapper.Parent;
            int i = templatesWrapper.Controls.IndexOf(templateWrapper);
            templatesWrapper.Controls.RemoveAt(i);
        }

        // 0.3. Копирование листов

        public void CopyLists(object sender, EventArgs e) 
        {
            WinForms.Button button = (WinForms.Button)sender;
            WinForms.FlowLayoutPanel optionsWrapper = (WinForms.FlowLayoutPanel)button.Parent;

            // Считываем параметры копирования
            WinForms.CheckBox stampCheck = (WinForms.CheckBox)optionsWrapper.Controls[0];
            WinForms.CheckBox viewsCheck = (WinForms.CheckBox)optionsWrapper.Controls[1];
            WinForms.CheckBox legendsCheck = (WinForms.CheckBox)optionsWrapper.Controls[2];
            WinForms.CheckBox specCheck = (WinForms.CheckBox)optionsWrapper.Controls[3];
            WinForms.CheckBox annoCheck = (WinForms.CheckBox)optionsWrapper.Controls[4];

            bool stampFlag = stampCheck.Checked;
            bool viewsFlag = viewsCheck.Checked;
            bool legendFlag = legendsCheck.Checked;
            bool specFlag = specCheck.Checked;
            bool annoFlag = annoCheck.Checked;

            // Находим форму
            WinForms.FlowLayoutPanel formWrapper = (WinForms.FlowLayoutPanel)optionsWrapper.Parent;
            WinForms.FlowLayoutPanel treeWrapper = (WinForms.FlowLayoutPanel)formWrapper.Controls[0];
            // Открываем транзакцию
            Transaction t = new Transaction(Doc, "Скопировать листы");
            t.Start();
            // Пробегаем по форме
            foreach (WinForms.FlowLayoutPanel buildingWrapper in treeWrapper.Controls)
            {
                WinForms.FlowLayoutPanel tomesWrapper = (WinForms.FlowLayoutPanel)buildingWrapper.Controls[1];

                foreach (WinForms.FlowLayoutPanel tomeWrapper in tomesWrapper.Controls)
                {
                    WinForms.FlowLayoutPanel sheetsWrapper = (WinForms.FlowLayoutPanel)tomeWrapper.Controls[1];

                    foreach (WinForms.FlowLayoutPanel sheetWrapper in sheetsWrapper.Controls)
                    {
                        SheetButton origSheetButton = (SheetButton)sheetWrapper.Controls[0];
                        ViewSheet originalSheet = origSheetButton.sheet;
                        WinForms.FlowLayoutPanel templatesWrapper = (WinForms.FlowLayoutPanel)sheetWrapper.Controls[1];

                        foreach (WinForms.FlowLayoutPanel templateWrapper in templatesWrapper.Controls)
                        {
                            System.Windows.Forms.TextBox buildingTB = (System.Windows.Forms.TextBox)templateWrapper.Controls[1];
                            System.Windows.Forms.TextBox tomeTB = (System.Windows.Forms.TextBox)templateWrapper.Controls[2];
                            System.Windows.Forms.TextBox numTB = (System.Windows.Forms.TextBox)templateWrapper.Controls[3];
                            System.Windows.Forms.TextBox nameTB = (System.Windows.Forms.TextBox)templateWrapper.Controls[4];

                            string buildingName = buildingTB.Text;
                            string tomeName = tomeTB.Text;
                            string num = numTB.Text;
                            string name = nameTB.Text;



                            if (buildingName == "<ПРИМЕЧАНИЕ НЕ ЗАДАНО>") buildingName = "";
                            if (tomeName == "<РАЗДЕЛ НЕ ЗАДАН>") tomeName = "";
                            int counter = 0; // Проверяем, уникален ли номер; если нет - создаем уникальный
                            string origNum = num;
                            while (!sheetNumbers.Add(num)) {
                                num = origNum + "$" + counter.ToString();
                                counter++;
                            }

                            // Начинаем копирование
                            // Забираем titleblock исходного листа
                            FamilyInstance titleBlock = new FilteredElementCollector(Doc, originalSheet.Id).
                                            OfCategory(BuiltInCategory.OST_TitleBlocks).
                                            FirstElement() as FamilyInstance;
                            // Создаем лист
                            ViewSheet newSheet = ViewSheet.Create(Doc, titleBlock.GetTypeId());
                            // Прописываем имя и номер
                            newSheet.Name = Regex.Replace(name, @"[\~#%&*{}/:<>?""]", string.Empty); 
                            newSheet.SheetNumber = num;

                            //Прописываем параметр тома
                            Autodesk.Revit.DB.Parameter paramTome = newSheet.LookupParameter("ADSK_Штамп Раздел проекта");
                            if (paramTome != null)
                            {
                                paramTome.Set(tomeName);
                            }
                            else {
                                paramTome = newSheet.LookupParameter("Раздел проекта");
                                if (paramTome != null)
                                {
                                    paramTome.Set(tomeName);
                                }
                            }
                            //Прописываем параметр здания
                            Autodesk.Revit.DB.Parameter paramBuilding = newSheet.LookupParameter("ADSK_Примечание");
                            if (paramBuilding != null)
                            {
                                paramBuilding.Set(buildingName);
                            }
                            else
                            {
                                paramBuilding = newSheet.LookupParameter("Примечание");
                                if (paramBuilding != null)
                                {
                                    paramBuilding.Set(buildingName);
                                }
                            }

                            //Переносим с исходного листа выбранные элементы
                            if (viewsFlag | legendFlag | specFlag | annoFlag)
                            {
                                HashSet<ElementId> copiedAndSkippedElementIds = new HashSet<ElementId>(); // Множество, куда будем записывать все скопированные и пропущенные элементы
                                
                                // Собираем все вьюпорты на исходном листе
                                ICollection<Element> viewports = new FilteredElementCollector(Doc, originalSheet.Id).
                                                                    OfClass(typeof(Viewport)).
                                                                    WhereElementIsNotElementType().
                                                                    ToElements();
                                // 1. Копирование видов и легенд
                                // Перебираем все виды на исходном листе (планы + разрезы + фасады + легенды)
                                foreach (ElementId elemId in originalSheet.GetAllPlacedViews())
                                {
                                    Autodesk.Revit.DB.View originalView = Doc.GetElement(elemId) as Autodesk.Revit.DB.View; // Вид, который мы копируем
                                    Autodesk.Revit.DB.View newView = null;                                                  // Новый вид

                                    // Проверяем, легенда этот вид или план/разрез/фасад
                                    if (originalView.ViewType == ViewType.Legend)
                                    {
                                        if (legendFlag)
                                        {
                                            newView = originalView;
                                        }
                                    }
                                    else // Если не легенда - копируем вид с детализацией
                                    {
                                        if (originalView.CanViewBeDuplicated(ViewDuplicateOption.WithDetailing))
                                        {
                                            if (viewsFlag)
                                            {
                                                ElementId newViewId = originalView.Duplicate(ViewDuplicateOption.WithDetailing);
                                                newView = Doc.GetElement(newViewId) as Autodesk.Revit.DB.View;
                                                string newName = originalView.Name;
                                                int tempCount = 0;
                                                while (!viewNames.Add(newName))
                                                {
                                                    newName += "$" + tempCount.ToString();
                                                    tempCount++;
                                                }
                                                newView.Name = newName;
                                            }
                                        }
                                    }

                                    foreach (Viewport viewport in viewports)
                                    {
                                        if (newView != null)
                                        {
                                            // Ищем вьюпорт вида, который хотим скопировать
                                            if (viewport.SheetId == originalSheet.Id && viewport.ViewId == originalView.Id)
                                            {
                                                XYZ center = viewport.GetBoxCenter();                                          // Находим центр вьюпорта
                                                Viewport newViewport = Viewport.Create(Doc, newSheet.Id, newView.Id, center); // Создаем новый вьюпорт в документе, на новом листе, на новом виде, с тем же расположением
                                            }
                                        }
                                        copiedAndSkippedElementIds.Add(viewport.Id); //Добавляем вид в список уже скопированных элементов
                                    }

                                    copiedAndSkippedElementIds.Add(elemId);
                                } // Конец перебора видов на исходном плане

                                // 2. Копирование спецификаций
                                // Собираем спецификации на листе
                                ICollection<Element> specs = new FilteredElementCollector(Doc, originalSheet.Id).
                                                            OfClass(typeof(ScheduleSheetInstance)).
                                                            WhereElementIsNotElementType().
                                                            ToElements();
                                // Собираем вьюпорты спецификаций (в проекте)
                                ICollection<Element> viewSpecs = new FilteredElementCollector(Doc).
                                                                OfClass(typeof(ViewSchedule)).
                                                                WhereElementIsNotElementType().
                                                                ToElements();
                                // Пробегаем по всем спецификациям
                                    foreach (ScheduleSheetInstance spec in specs)
                                    {
                                        if (specFlag)
                                        {
                                            // Проверяем, что спецификация - не ревизия внутри тайтлблока
                                            if (!spec.IsTitleblockRevisionSchedule)
                                            {
                                                foreach (ViewSchedule viewSpec in viewSpecs)
                                                {
                                                    if (spec.ScheduleId == viewSpec.Id)
                                                    {
                                                        XYZ center = spec.Point;
                                                        ScheduleSheetInstance newSpec = ScheduleSheetInstance.Create(Doc, newSheet.Id, viewSpec.Id, center);
                                                    }
                                                    // Добавляем элемент в список скопированных
                                                }
                                            }
                                        }
                                        copiedAndSkippedElementIds.Add(spec.Id);
                                    }
                                // 3. Копирование остальных аннотативных элементов
                                if (annoFlag)
                                {
                                    IEnumerable<ElementId> elementsOnSheetIds = new FilteredElementCollector(Doc, originalSheet.Id).ToElementIds();
                                    var annotationElementsId = new List<ElementId>();
                                    foreach (ElementId eId in elementsOnSheetIds)
                                    {
                                        if (!copiedAndSkippedElementIds.Contains(eId))
                                        {
                                            annotationElementsId.Add(eId);                                            
                                        }
                                    }
                                    ElementTransformUtils.CopyElements(originalSheet, annotationElementsId, newSheet, null, null);
                                }                               

                            } // Конец копирования опциональных элементов

                            // 4. Копирование данных штампа
                            if (stampFlag)
                            {
                                HashSet<string> paramNames = new HashSet<string>();

                                string paramName = "ADSK_Штамп Наименование объекта";
                                paramNames.Add(paramName);
                                paramName = "ADSK_Штамп Количество листов";
                                paramNames.Add(paramName);
                                paramName = "ADSK_Штамп Инвентарный номер";
                                paramNames.Add(paramName);

                                for (int i = 1; i <= 6; i++) 
                                {
                                    paramName = "ADSK_Штамп Боковой Строка " + i.ToString() + " должность";
                                    paramNames.Add(paramName);

                                    paramName = "ADSK_Штамп Боковой Строка " + i.ToString() + " фамилия";
                                    paramNames.Add(paramName);

                                    paramName = "ADSK_Штамп Строка " + i.ToString() + " должность";
                                    paramNames.Add(paramName);

                                    paramName = "ADSK_Штамп Строка " + i.ToString() + " фамилия";
                                    paramNames.Add(paramName);
                                }

                                foreach (string paramN in  paramNames) 
                                {
                                    Parameter origParam = originalSheet.LookupParameter(paramN);
                                    Parameter newParam = newSheet.LookupParameter(paramN);

                                    if ((origParam != null) && (newParam != null))
                                    {
                                        bool flag = newParam.Set(origParam.AsString());
                                    }
                                }
                            } // Конец копирования данных штампа
                        } // Конец перебора шаблонов
                    } // Конец перебора листов
                } // Конец перебора томов
            } // Конец перебора зданий
            t.Commit();
            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
        } // Конец события CopyLists
    }

    //-------------------------------------------
    // 1. Классы
    //-------------------------------------------

    // 1.1. Класс кнопки с существующим листом
    public class SheetButton : WinForms.Button
    {
        public ViewSheet sheet;
        public string buildingName;
        public string tomeName;
    }
}
