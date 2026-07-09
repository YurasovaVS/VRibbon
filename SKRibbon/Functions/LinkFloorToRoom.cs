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
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

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

            // Выделение 
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Selection selection = uidoc.Selection;
            ICollection<ElementId> selectionIds = uidoc.Selection.GetElementIds();

            if (selectionIds.Count <= 0)
            {
                TaskDialog.Show("Выделение", "Вы ничего не выделили");
                return Result.Cancelled;

            }

            // ---------------------------------------------

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

            foreach (ElementId floorId in selectionIds)
            {
                bool flag = false;
                Element floor = doc.GetElement(floorId);
                if (floor.Name.Contains("КЖ_М")) continue;
                BoundingBoxXYZ boundingBox = floor.get_BoundingBox(null);

                IList<Reference> topFaceRefs = HostObjectUtils.GetTopFaces(floor as Floor);
                if (topFaceRefs.Count == 0) continue;
                Face topFace = floor.GetGeometryObjectFromReference(topFaceRefs[0]) as Face;
                if (topFace == null) continue;


                XYZ point = new XYZ((boundingBox.Max.X + boundingBox.Min.X)/2, (boundingBox.Max.Y + boundingBox.Min.Y) / 2, boundingBox.Max.Z); // в z раньше было +2, добавить перед проверкой на комнату

                // Проецируем точку на грань
                IntersectionResult projectResult = topFace.Project(point);
                if (projectResult != null)
                {
                    UV uvPoint = projectResult.UVPoint;
                    // Verify if the projected point lies within the physical boundary loops
                    if (topFace.IsInside(uvPoint))
                    {
                        flag = true;
                    }
                }

                // Если центральная точка не работает, проходимся по всем точкам
                if (!flag)
                {
                    point = new XYZ(boundingBox.Min.X, boundingBox.Min.Y, boundingBox.Max.Z);
                    do
                    {
                        point = new XYZ(point.X + 1, point.Y, point.Z);
                        do
                        {
                            point = new XYZ(point.X, point.Y + 1, point.Z);

                            projectResult = topFace.Project(point);
                            if (projectResult != null)
                            {
                                UV uvPoint = projectResult.UVPoint;
                                if (topFace.IsInside(uvPoint))
                                {
                                    flag = true;
                                }
                            }
                        } while ((point.X < boundingBox.Max.X) && (point.Y < boundingBox.Max.Y) && (flag == false));
                    } while ((point.X < boundingBox.Max.X) && (point.Y < boundingBox.Max.Y) && (flag == false));
                }

                // Если точка на полу найдена, проверяем, попадает ли точка над ней в комнату
                if (flag)
                {
                    point = new XYZ(point.X, point.Y, point.Z + 2);
                    foreach (Element room in rooms)
                    {
                        Room trueRoom = room as Room;
                        if (trueRoom.IsPointInRoom(point))
                        {
                            Parameter groupParam = floor.LookupParameter("ADSK_Номер помещения квартиры");
                            if (groupParam != null) groupParam.Set(trueRoom.Number.ToString());
                            break;
                        }
                    }
                }
                                
            }
            t.Commit();

            return Result.Succeeded;
            //throw new NotImplementedException();
        }
    }
}
