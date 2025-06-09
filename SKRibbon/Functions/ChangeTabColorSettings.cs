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
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CopyListsTree;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    internal class ChangeTabColorSettings : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            using (VForm form = new ChangeColorSettingsForm(uiApp))
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }
        }
    }
}
