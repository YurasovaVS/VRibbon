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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections;
using System.Windows.Forms;

namespace SKRibbon
{
    public class FormUtils
    {
        //   1 - Собираем все листы в основу для древа

        /* Словарь Зданий
         *      Здание : Словарь Томов
         *          Том : Список объектов
         *              Объект ViewSheet
         *                  
        */
        public static Dictionary<string, Dictionary<string, List<ViewSheet>>> CollectSheetDictionary (Document doc, bool orderByNumber)
        {
            Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();

            ICollection<Element> sheets = new FilteredElementCollector(doc).
                                            OfCategory(BuiltInCategory.OST_Sheets).
                                            WhereElementIsNotElementType().
                                            ToElements();
            foreach (ViewSheet sheet in sheets)
            {
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
                        List<ViewSheet> sheetsList = new List<ViewSheet>();
                        buildingsDict[building].Add(tome, sheetsList);
                    }
                    // Добавляем лист в нужный том
                    buildingsDict[building][tome].Add(sheet);
                }
            } // Конец создания словаря листов

            if (orderByNumber)
            {
                Dictionary<string, Dictionary<string, List<ViewSheet>>> tempDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();

                foreach (var building in buildingsDict)
                {
                    Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                    tempDict.Add(building.Key, tomesDict);
                    foreach (var tome in building.Value)
                    {
                        List<ViewSheet> sheetsList = tome.Value.OrderBy(sheet => sheet.SheetNumber).ToList();
                        tempDict[building.Key].Add(tome.Key, sheetsList);
                    }
                }

                buildingsDict = tempDict;
            }

            return buildingsDict;
        }

        public static WinForms.TreeView CreateSheetTreeView (Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict)
        {
            WinForms.TreeView tree = new WinForms.TreeView ();
            foreach (var building in buildingsDict)
            {
                tree.Nodes.Add(building.Key);
                Dictionary<string, List<ViewSheet>> tomes = building.Value;

                foreach (var tome in tomes)
                {
                    int buildingIndex = tree.Nodes.Count - 1;
                    tree.Nodes[buildingIndex].Nodes.Add(tome.Key);
                    List<ViewSheet> treeSheets = tome.Value;
                    treeSheets = treeSheets.OrderBy(sheet => sheet.SheetNumber).ToList();
                    //orderBy

                    foreach (var sheet in treeSheets)
                    {
                        SheetTreeNode node = new SheetTreeNode();
                        node.Text = sheet.SheetNumber + "   |   " + sheet.Name;
                        node.sheet = sheet;
                        int tomeIndex = tree.Nodes[buildingIndex].Nodes.Count - 1;
                        tree.Nodes[buildingIndex].Nodes[tomeIndex].Nodes.Add(node);
                    }
                }
            } // Конец построения дерева
                        
            return tree;
        }

        // Класс конечных нодов дерева, содержащих ссылки на листы
        public class SheetTreeNode : WinForms.TreeNode
        {
            public ViewSheet sheet;
        }


    }
}
