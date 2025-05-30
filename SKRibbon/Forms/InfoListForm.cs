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
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace InfoListMaker
{
    public partial class InfoListForm : System.Windows.Forms.Form
    {
        Document Doc;
        string SignaturesPath = @"\\absknas\переезд\13_Пользователи\_ПодписиСК";
        string SavePath = "";
        System.Windows.Forms.TextBox pathTextBox = new System.Windows.Forms.TextBox();
        Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        public InfoListForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;

            SavePath = Path.GetDirectoryName(Doc.PathName);
            if ((SavePath == null) || (SavePath == ""))
            {
                SavePath = "C:\\";
            }
            if (Directory.Exists(SavePath))
            {
                SavePath += "infoLists";
                if (!Directory.Exists(SavePath))
                {
                    Directory.CreateDirectory(SavePath);
                }
            }
            string sp = SavePath;

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.Padding = new Padding(10, 10, 10, 10);

            // Инициализация формы
            // 0. Заголовок выбора папки с подписями
            System.Windows.Forms.Label headerSign = new System.Windows.Forms.Label();
            headerSign.Size = new System.Drawing.Size(300, 30);
            headerSign.Margin = new Padding(0, 5, 0, 0);
            headerSign.Anchor = AnchorStyles.Left;
            headerSign.Text = "Выберите папку с подписями:";

            // 1. Строка выбора шаблона
            // Обертка
            FlowLayoutPanel pathTextBoxWrapper = new FlowLayoutPanel();
            pathTextBoxWrapper.FlowDirection = FlowDirection.LeftToRight;
            pathTextBoxWrapper.AutoSize = true;
            pathTextBoxWrapper.Padding = new Padding(0, 0, 0, 0);

            // 1.1. Строка (текстбокс)
            // переехал в глобальные параметры
            pathTextBox.Size = new System.Drawing.Size(400, 30);
            pathTextBox.Anchor = AnchorStyles.Left;
            pathTextBox.Text = SignaturesPath;

            // 1.2. Кнопка выбора файла
            System.Windows.Forms.Button chooseFolderButton = new System.Windows.Forms.Button();
            chooseFolderButton.Size = new System.Drawing.Size(30, 30);
            chooseFolderButton.Anchor = AnchorStyles.Left;
            chooseFolderButton.Click += ChooseFolder;

            // Добавление элементов в обертку
            pathTextBox.Parent = pathTextBoxWrapper;
            chooseFolderButton.Parent = pathTextBoxWrapper;

            pathTextBoxWrapper.Controls.Add(pathTextBox);
            pathTextBoxWrapper.Controls.Add(chooseFolderButton);

            // 2.0. Заголовок шифра
            System.Windows.Forms.Label headerId = new System.Windows.Forms.Label();
            headerId.Size = new System.Drawing.Size(300, 30);
            headerId.Margin = new Padding(0, 5, 0, 0);
            headerId.Anchor = AnchorStyles.Left;
            headerId.Text = "Введите шифр:";
            
            // 2. Строка с шифром
            System.Windows.Forms.TextBox projectNumTB = new System.Windows.Forms.TextBox();
            projectNumTB.Size = new System.Drawing.Size(400, 30);
            projectNumTB.Anchor = AnchorStyles.Left;

            // 3. Список томов
            // Собираем все тома, и относящиеся к ним листы
            ICollection<Element> sheets = new FilteredElementCollector(Doc).
                                            OfCategory(BuiltInCategory.OST_Sheets).
                                            WhereElementIsNotElementType().
                                            ToElements();

            foreach (ViewSheet sheet in sheets)
            {
                Autodesk.Revit.DB.Parameter tomeParam = sheet.LookupParameter("ADSK_Штамп Раздел проекта");
                Autodesk.Revit.DB.Parameter buildingParam = sheet.LookupParameter("ADSK_Примечание");
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
                    string key = building + " - " + tome;
                    
                    if (tome != "<РАЗДЕЛ НЕ ЗАДАН>")
                    {
                        // Если в словаре нет такого тома, создаем его
                        if (!tomesDict.ContainsKey(key))
                        {
                            List<ViewSheet> sheetList = new List<ViewSheet>();
                            tomesDict.Add(key, sheetList);
                        }
                        // Добавляем лист в нужный том
                        tomesDict[key].Add(sheet);
                    }                    
                }
            }

            CheckedListBox tomesCheckList = new CheckedListBox();
            tomesCheckList.Size = new System.Drawing.Size(400, 300);

            foreach (var tome in tomesDict) 
            {
                    tomesCheckList.Items.Add(tome.Key);
            }

            // 4. Кнопка запуска программы
            System.Windows.Forms.Button button = new System.Windows.Forms.Button();
            button.Size = new Size(200, 50);
            button.Anchor = AnchorStyles.Top;
            button.Text = "Собрать ИУЛы";
            button.Click += RunIulMaker;

            headerSign.Parent = formWrapper;
            pathTextBoxWrapper.Parent = formWrapper;
            headerId.Parent = formWrapper;
            projectNumTB.Parent = formWrapper;
            tomesCheckList.Parent = formWrapper;
            button.Parent = formWrapper;

            formWrapper.Controls.Add(headerSign);
            formWrapper.Controls.Add(pathTextBoxWrapper);
            formWrapper.Controls.Add(headerId);
            formWrapper.Controls.Add(projectNumTB);
            formWrapper.Controls.Add(tomesCheckList);
            formWrapper.Controls.Add(button);

            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);
            this.Size = new System.Drawing.Size(500, 600);
        }

        // ---------------------- Функции ------------------------------------
        public void RunIulMaker (object sender, EventArgs e)
        {
            System.Windows.Forms.Button button = (System.Windows.Forms.Button)sender;
            FlowLayoutPanel pathWrapper = (FlowLayoutPanel)formWrapper.Controls[1];
            System.Windows.Forms.TextBox projectNumTB = (System.Windows.Forms.TextBox)formWrapper.Controls[3];
            CheckedListBox tomesCheckList = (CheckedListBox)formWrapper.Controls[4];
            foreach (string key in tomesCheckList.CheckedItems) { 
                Dictionary<string, HashSet<string>> posNamePairs = new Dictionary<string, HashSet<string>>();
                StringBuilder sb = new StringBuilder();
                foreach (ViewSheet sheet in tomesDict[key])
                {
                    
                    // Пробегаемся по всем парам Должность - Имя
                    for (int index = 1; index <= 6; index++)
                    {
                        string posParamName = "ADSK_Штамп Строка " + (index) + " должность";
                        string nameParamName = "ADSK_Штамп Строка " + (index) + " фамилия";

                        string posParamValue = sheet.LookupParameter(posParamName).AsString();
                        string nameParamValue = sheet.LookupParameter(nameParamName).AsString();


                        if (!String.IsNullOrEmpty(posParamValue) && !String.IsNullOrEmpty(nameParamValue))
                        {
                            if (!posNamePairs.ContainsKey(posParamValue)) {
                                HashSet<string> hashSet = new HashSet<string>();
                                posNamePairs.Add(posParamValue, hashSet);
                            }
                            posNamePairs[posParamValue].Add(nameParamValue);
                        }
                    }
                }
                
                // Открываем Excel
                Excel.Application eApp = new Excel.Application
                {
                    // Отображаем окно
                    Visible = true,
                    // Листов в рабочей книге
                    SheetsInNewWorkbook = 1
                };
                // Добавляем рабочую книгу
                Excel.Workbook workBook = eApp.Workbooks.Add(Type.Missing);
                // Отключаем отображение всплывающих окон
                eApp.DisplayAlerts = false;
                Excel.Worksheet wSheet = (Excel.Worksheet)eApp.Worksheets.get_Item(1);
                wSheet.Name = "Sheet1";

                // Заполняем первую строку заголовков

                Excel.Range rBegin = (Excel.Range)wSheet.Cells[1, 1];
                Excel.Range r2 = (Excel.Range)wSheet.Cells[1, 2];
                Excel.Range r3 = (Excel.Range)wSheet.Cells[1, 3];
                Excel.Range r4 = (Excel.Range)wSheet.Cells[1, 4];
                Excel.Range rEnd = (Excel.Range)wSheet.Cells[1, 5];

                rBegin.EntireColumn.ColumnWidth = 12.71;
                r2.EntireColumn.ColumnWidth = 30;
                r3.EntireColumn.ColumnWidth = 30;
                r4.EntireColumn.ColumnWidth = 10;
                rEnd.EntireColumn.ColumnWidth = 17.3;

                // Заполняем 1 строку
                FillExcelRow(wSheet, 1, "Номер п/п", "Обозначение документа (шифр)", "Наименование документа", "Версия", "Номер последнего изменения", true);                
                // Заполняем 2 строку
                string tomeId = (projectNumTB.Text != "") ? projectNumTB.Text + "-" : "";
                tomeId += Regex.Replace(key, " ", string.Empty);
                FillExcelRow(wSheet, 2, "1", tomeId, "", "1", "1", false);

                // Заполняем пятую строку
                wSheet.Cells[5, 1] = "MD5";
                wSheet.Cells[5, 2] = "1";
                wSheet.Cells[5, 3] = "";
                wSheet.Cells[5, 4] = "";
                wSheet.Cells[5, 5] = "";

                rBegin = (Excel.Range)wSheet.Cells[5, 1];
                r2 = (Excel.Range)wSheet.Cells[5, 2];
                r3 = (Excel.Range)wSheet.Cells[5, 3];
                r4 = (Excel.Range)wSheet.Cells[5, 4];
                rEnd = (Excel.Range)wSheet.Cells[5, 5];

                Excel.Range row5Range = wSheet.get_Range(rBegin, rEnd);
                SetRangeParams(row5Range, false);

                row5Range = wSheet.get_Range(rBegin, rBegin);
                row5Range.Cells.Font.Bold = true;

                row5Range = wSheet.get_Range(r3, rEnd);
                row5Range.Merge(Type.Missing);

                // Заполняем 8 строку
                FillExcelRow(wSheet, 8, "Номер п/п", "Наименование файла", "Дата и время последнего изменения файла", "Размер файла (байт)", true);
                // Заполняем 9 строку
                FillExcelRow(wSheet, 9, "1", tomeId + ".pdf", "", "", false);
                // Заполняем 12 строку
                FillExcelRow(wSheet, 12, "Характер работы", "ФИО", "Подпись", "Дата подписания", true);
                // Заполняем 13 строку

                rBegin = (Excel.Range)wSheet.Cells[13, 1];
                rEnd = (Excel.Range)wSheet.Cells[13, 5];
                Excel.Range row13Range = wSheet.get_Range(rBegin, rEnd);
                SetRangeParams(row13Range, false);

                row13Range.Merge(Type.Missing);
                wSheet.Cells[13, 1] = "1";
                // В цикле заполняем строки с работниками
                int i = 14;
                foreach (var position in posNamePairs)
                {
                    foreach (string name in position.Value)
                    {
                        FillExcelRow(wSheet, i, 48, position.Key, name);
                        i++;
                    }
                }
                Regex pattern = new Regex("[;<>,!.+= ]");
                string filename = "УЛ-УЛ-" + pattern.Replace(key, "") + ".xlsx";
                string fileSavePath = Path.Combine(SavePath, filename);
                if (File.Exists(fileSavePath))
                {
                    File.Delete(fileSavePath);
                }
                eApp.Application.ActiveWorkbook.SaveAs(fileSavePath, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                eApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(eApp);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public void SetRangeParams (Excel.Range range, bool isBold)
        {
            range.Cells.Font.Name = "Times New Roman";
            range.Cells.Font.Size = 11;
            range.Cells.Font.Bold = isBold;
            range.Cells.Borders.Color = ColorTranslator.ToOle(System.Drawing.Color.Black);
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Cells.WrapText = true;
        }
        // Заполнение строк (перегрузки функции FillExcelRow)
        public void FillExcelRow(Excel.Worksheet wSheet, int row, string cell_1, string cell_2, string cell_3, string cell_4, string cell_5, bool isHeader)
        {
            Excel.Range rBegin = (Excel.Range)wSheet.Cells[row, 1];
            Excel.Range rEnd = (Excel.Range)wSheet.Cells[row, 5];
            
            Excel.Range range = wSheet.get_Range(rBegin, rEnd);
            SetRangeParams(range, isHeader);

            wSheet.Cells[row, 1] = cell_1;
            wSheet.Cells[row, 2] = cell_2;
            wSheet.Cells[row, 3] = cell_3;
            wSheet.Cells[row, 4] = cell_4;
            wSheet.Cells[row, 5] = cell_5;
        }

        public void FillExcelRow(Excel.Worksheet wSheet, int row, string cell_1, string cell_2, string cell_3, string cell_4, bool isHeader)
        {
            Excel.Range rBegin = (Excel.Range)wSheet.Cells[row, 1];
            Excel.Range r4 = (Excel.Range)wSheet.Cells[row, 4];
            Excel.Range rEnd = (Excel.Range)wSheet.Cells[row, 5];

            Excel.Range range = wSheet.get_Range(rBegin, rEnd);
            SetRangeParams(range, isHeader);

            range = wSheet.get_Range(r4, rEnd);
            range.Merge(Type.Missing);

            wSheet.Cells[row, 1] = cell_1;
            wSheet.Cells[row, 2] = cell_2;
            wSheet.Cells[row, 3] = cell_3;
            wSheet.Cells[row, 4] = cell_4;
        }

        public void FillExcelRow(Excel.Worksheet wSheet, int row, float height, string position, string name)
        {
            Excel.Range rBegin = (Excel.Range)wSheet.Cells[row, 1];
            Excel.Range r2 = (Excel.Range)wSheet.Cells[row, 2];
            Excel.Range r3 = (Excel.Range)wSheet.Cells[row, 3];
            Excel.Range r4 = (Excel.Range)wSheet.Cells[row, 4];
            Excel.Range rEnd = (Excel.Range)wSheet.Cells[row, 5];

            Excel.Range range = wSheet.get_Range(rBegin, rEnd);
            SetRangeParams(range, false);
            range.EntireRow.RowHeight = height;

            range = wSheet.get_Range(r4, rEnd);
            range.Merge(Type.Missing);

            wSheet.Cells[row, 1] = position;
            wSheet.Cells[row, 2] = name;
            // Если в папке с подписями лежит подпись...
            string filePath = SignaturesPath + "\\" + name + ".png";
            if (File.Exists(filePath)) {
                wSheet.Shapes.AddPicture(filePath, 
                    Microsoft.Office.Core.MsoTriState.msoFalse, 
                    Microsoft.Office.Core.MsoTriState.msoCTrue, 
                    (float)r3.Left + 5, (float)r3.Top + 2, -1, -1);
            }
        }

        // Событие вызова окна с выбором папки
        public void ChooseFolder(object sender, EventArgs e)
        {
            FlowLayoutPanel wrapper = (FlowLayoutPanel)formWrapper.Controls[1];
            System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)wrapper.Controls[0];

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox.Text = dialog.SelectedPath;
            }
        }
    }
}
