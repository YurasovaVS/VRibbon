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

using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Controls;
using Autodesk.Revit.DB.Architecture;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    class LinkFloorToRoom : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            // Временно
            Document Doc = doc;

            ICollection<Element> rooms = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Rooms).
                                                    WhereElementIsNotElementType().
                                                    ToElements();

            ICollection<Element> floors = new FilteredElementCollector(Doc).
                                                    OfCategory(BuiltInCategory.OST_Floors).
                                                    WhereElementIsNotElementType().
                                                    ToElements();

            Transaction t = new Transaction(Doc, "Связать полы с помещениями");
            t.Start();

            foreach (Element floor in floors)
            {
                if (floor.Name.Contains("КЖ_М")) continue;
                BoundingBoxXYZ boundingBox = floor.get_BoundingBox(null);
                
                XYZ point = new XYZ((boundingBox.Max.X + boundingBox.Min.X)/2, (boundingBox.Max.Y + boundingBox.Min.Y) / 2, boundingBox.Max.Z + 2);
                                
                foreach (Element room in rooms)
                {
                    Room trueRoom = room as Room;
                    if (trueRoom.IsPointInRoom(point))
                    {
                        Parameter groupParam = floor.LookupParameter("ADSK_Группирование");
                        if (groupParam != null) groupParam.Set(trueRoom.Number.ToString()); 
                        break;
                    }
                }                
            }
            t.Commit();

            return Result.Succeeded;
            //throw new NotImplementedException();
        }
    }
}
