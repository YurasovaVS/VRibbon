using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Drawing.Printing;

namespace BatchPrinting
{
    
    public partial class BatchPrintForm : System.Windows.Forms.Form
    {
        Document Doc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        string SavePath;


        public BatchPrintForm(Document doc)
        {
            // -------
            /*
            string printerName = "Adobe PDF";
            System.Drawing.Printing.PaperSize[] customPaperSizes = 
            { 
                new System.Drawing.Printing.PaperSize("A4", 0, 0),
                new System.Drawing.Printing.PaperSize("A3", 0, 0),
                new System.Drawing.Printing.PaperSize("A2", 0, 0),
                new System.Drawing.Printing.PaperSize("A1", 0, 0),
                new System.Drawing.Printing.PaperSize("A0", 0, 0),
                new System.Drawing.Printing.PaperSize("3xA3-test", 3507, 1654)
            };
            //bool bGotPrinter = OpenPrinter(printerName, out hPrinter, ref defaults);

            PrintDocument pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;
            
            foreach (var paperSize in pd.PrinterSettings.PaperSizes)
            {
                Debug.WriteLine(paperSize.ToString());
            }
            
            foreach (System.Drawing.Printing.PaperSize customPaperSize in customPaperSizes)
            {
                bool flag = false;
                foreach (System.Drawing.Printing.PaperSize paperSize in pd.PrinterSettings.PaperSizes)
                {
                    if (customPaperSize.PaperName == paperSize.PaperName) flag = true;
                }              
                if (!flag)
                {
                    pd.PrinterSettings.PaperSizes.Add(customPaperSize);
                }                
            }
            

            foreach (var paperSize in pd.PrinterSettings.PaperSizes)
            {
                Debug.WriteLine(paperSize.ToString());
            }
            */
            // -------

            InitializeComponent();
            Doc = doc;
            SavePath = Path.GetDirectoryName(doc.PathName);
            if ((SavePath == null) || (SavePath == ""))
            {
                SavePath = @"C:";
            }
            if (Directory.Exists(SavePath))
            {
                SavePath += @"\pdf";
                if (!Directory.Exists(SavePath))
                {
                    Directory.CreateDirectory(SavePath);
                }                
            }
            this.AutoScroll = true;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

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
                Parameter tomeParam = sheet.LookupParameter("ADSK_Штамп Раздел проекта");
                Parameter buildingParam = sheet.LookupParameter("ADSK_Примечание");
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

            // Создаем дерево листов проекта
            TreeView tree = new TreeView();
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
            // Инициализируем параметры дерева и добавляем его в форму
            tree.CheckBoxes = true;            
            tree.AfterCheck += node_AfterCheck;
            tree.Width = 300;
            tree.Height = 400;
            tree.Margin = new Padding(0, 10, 0, 0);
            tree.Parent = formWrapper;
            formWrapper.Controls.Add(tree);
            tree.Anchor = AnchorStyles.Left;

            // Добавляем wrapper для опций
            // 
            FlowLayoutPanel optionsWrapper = new FlowLayoutPanel();
            optionsWrapper.AutoSize = true;
            optionsWrapper.Anchor = AnchorStyles.Top;
            optionsWrapper.FlowDirection = FlowDirection.TopDown;
            optionsWrapper.Margin = new Padding(30, 10, 0, 0);
            optionsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(optionsWrapper);

            // Добавляем заголовок выпадающего списка
            Label printersHeader = new Label();
            printersHeader.Size = new Size(300, 30);
            printersHeader.Margin = new Padding(0, 30, 0, 0);
            printersHeader.Text = "Выберите принтер:";
            printersHeader.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(printersHeader);

            // Добавляем выпадающий список со списком принтеров
            System.Windows.Forms.ComboBox printersListCB = new System.Windows.Forms.ComboBox();
            printersListCB.Size = new Size(300, 30);
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                if ((printer == "Adobe PDF") ||
                    (printer == "Microsoft Print to PDF")
                    )
                printersListCB.Items.Add(printer);
            }
            printersListCB.SelectedIndex = 0;
            printersListCB.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(printersListCB);

            // Добавляем путь
            Label pathHeader = new Label();
            pathHeader.Size = new Size(300, 50);
            pathHeader.Margin = new Padding(0, 30, 0, 0);
            pathHeader.Text = "Файлы будут сохранены в папку:\n" + SavePath;
            pathHeader.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(pathHeader);

            // Добавляем кнопку
            Button okButton = new Button();
            okButton.Text = "Вывести листы";
            okButton.Size = new Size(100, 60);
            okButton.Click += PrintSheets;
            okButton.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(okButton);
            okButton.Anchor = AnchorStyles.Left;
        }

        // --------------
        // Классы
        // --------------

        private class SheetTreeNode : System.Windows.Forms.TreeNode
        {
            public ViewSheet sheet;
        }


        // --------------
        // События и методы
        // --------------

        // Проставление галочек напротив всех "детей"
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    // Если у детей нода тоже есть дети, рекурсивно вызываем для них функцию
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        private void node_AfterCheck(object sender, TreeViewEventArgs e)
        {            
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }

        // Обработка дерева и вывод листов
        private void PrintSheets(object sender, EventArgs e)
        {
            TreeView tree = (TreeView)formWrapper.Controls[0];
            FlowLayoutPanel optionsWrapper = (FlowLayoutPanel)formWrapper.Controls[1];
            System.Windows.Forms.ComboBox printersCB = (System.Windows.Forms.ComboBox)optionsWrapper.Controls[1];
            //CheckBox multiPagesCheck = (CheckBox)optionsWrapper.Controls[4];
            string saveAsDialogCaption = "";

            switch (printersCB.SelectedItem.ToString())
            {
                case "Adobe PDF":
                    saveAsDialogCaption = "Сохранить PDF-файл как";
                    break;
                case "Microsoft Print to PDF":
                    saveAsDialogCaption = "Сохранение результата печати";
                    break;
            }

            Transaction t = new Transaction(Doc, "Напечатать листы");
            t.Start();

            foreach (TreeNode building in tree.Nodes) 
            {                
                foreach (TreeNode tome in building.Nodes)
                {
                    
                    foreach (SheetTreeNode sheet in tome.Nodes)
                    {
                        if (sheet.Checked)
                        {
                            FamilyInstance sheetInstance = new FilteredElementCollector(Doc, sheet.sheet.Id).
                                                            OfClass(typeof(FamilyInstance)).
                                                            OfCategory(BuiltInCategory.OST_TitleBlocks).
                                                            FirstOrDefault() as FamilyInstance;
                            if (sheetInstance != null)
                            {
                                ElementId sheetTypeId = sheetInstance.GetTypeId();

                                Parameter sheetHeight = sheetInstance.LookupParameter("Высота_Реальная");                                
                                Parameter sheetWidth = sheetInstance.LookupParameter("Ширина_Реальная");
                                if (sheetHeight == null)
                                {
                                    sheetHeight = sheetInstance.LookupParameter("Высота листа");
                                }
                                if (sheetWidth == null)
                                {
                                    sheetWidth = sheetInstance.LookupParameter("Ширина листа");
                                }

                                if ((sheetHeight == null) || (sheetWidth == null))
                                {
                                    TaskDialog.Show("Ошибка", "Не получилось определить размеры листа " + sheet.sheet.SheetNumber + " - " + sheet.sheet.Name);
                                }
                                else                                 
                                {
                                    double sheetWidthMM = UnitUtils.ConvertFromInternalUnits(sheetWidth.AsDouble(), UnitTypeId.Millimeters);
                                    double sheetHeightMM = UnitUtils.ConvertFromInternalUnits(sheetHeight.AsDouble(), UnitTypeId.Millimeters);

                                    var printManager = Doc.PrintManager;
                                    printManager.PrintSetup.CurrentPrintSetting = printManager.PrintSetup.InSession;
                                    printManager.SelectNewPrintDriver(printersCB.SelectedItem.ToString());
                                    printManager.PrintSetup.SaveAs("NewPrint");

                                    FilteredElementCollector collector = new FilteredElementCollector(Doc).OfClass(typeof(PrintSetting));
                                    PrintSetting setting = null;
                                    foreach (PrintSetting printSetting in collector)
                                    {
                                        if (printSetting.Name == "NewPrint")
                                        {
                                            setting = printSetting;
                                            break;
                                        }
                                    }

                                    printManager.PrintSetup.CurrentPrintSetting = setting;
                                    printManager.Apply();

                                    // Задаем параметры печати (расположение листа, увеличение)
                                    PrintParameters printParam = printManager.PrintSetup.CurrentPrintSetting.PrintParameters;
                                    printParam.PaperPlacement = PaperPlacementType.Center;
                                    printParam.ZoomType = ZoomType.Zoom;
                                    printParam.Zoom = 100;

                                    //Ориентация листа
                                    if (sheetHeightMM > sheetWidthMM)
                                    {
                                        printParam.PageOrientation = PageOrientationType.Portrait;
                                    }
                                    else
                                    {
                                        printParam.PageOrientation = PageOrientationType.Landscape;
                                    }

                                    string errorMessage = "Нестандартный размер листа ";
                                    string currentPaperSize = "";
                                    // А4
                                    if (
                                        ((sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00) &&  // Если высота приблизительно равна 297
                                        (sheetWidthMM >= 209.00 && sheetWidthMM <= 211.00))       // Если ширина приблизительно равна 210
                                        ||                                                        // или
                                        ((sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00) &&    // Если ширина приблизительно равна 297
                                        (sheetHeightMM >= 209.00 && sheetHeightMM <= 211.00))     // Если высота приблизительно равна 210
                                        )
                                    {
                                        currentPaperSize = "A4";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    // А3
                                    if (
                                        ((sheetHeightMM >= 419.00 && sheetHeightMM <= 421.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 419.00 && sheetWidthMM <= 421.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "A3";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //А2
                                    if (
                                        ((sheetHeightMM >= 419.00 && sheetHeightMM <= 421.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 593.00 && sheetWidthMM <= 595.00))       // Если ширина приблизительно равна 594
                                        ||                                                        // или
                                        ((sheetWidthMM >= 419.00 && sheetWidthMM <= 421.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 593.00 && sheetHeightMM <= 595.00))     // Если высота приблизительно равна 594
                                        )
                                    {
                                        currentPaperSize = "A2";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //А1
                                    if (
                                        ((sheetHeightMM >= 840.00 && sheetHeightMM <= 842.00) &&  // Если высота приблизительно равна 841
                                        (sheetWidthMM >= 593.00 && sheetWidthMM <= 595.00))       // Если ширина приблизительно равна 594
                                        ||                                                        // или
                                        ((sheetWidthMM >= 840.00 && sheetWidthMM <= 842.00) &&    // Если ширина приблизительно равна 841
                                        (sheetHeightMM >= 593.00 && sheetHeightMM <= 595.00))     // Если высота приблизительно равна 594
                                        )
                                    {
                                        currentPaperSize = "A1";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //А0
                                    if (
                                        ((sheetHeightMM >= 840.00 && sheetHeightMM <= 842.00) &&  // Если высота приблизительно равна 841
                                        (sheetWidthMM >= 1188.00 && sheetWidthMM <= 1190.00))       // Если ширина приблизительно равна 594
                                        ||                                                        // или
                                        ((sheetWidthMM >= 840.00 && sheetWidthMM <= 842.00) &&    // Если ширина приблизительно равна 841
                                        (sheetHeightMM >= 1188.00 && sheetHeightMM <= 1190.00))     // Если высота приблизительно равна 594
                                        )
                                    {
                                        currentPaperSize = "A0";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //3xA4
                                    if (
                                        ((sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 629.00 && sheetWidthMM <= 631.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 629.00 && sheetHeightMM <= 631.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "3xA4";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //4xA4
                                    if (
                                        ((sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 839.00 && sheetWidthMM <= 841.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 839.00 && sheetHeightMM <= 841.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "4xA4";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //5xA4
                                    if (
                                        ((sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1049.00 && sheetWidthMM <= 1051.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1049.00 && sheetHeightMM <= 1051.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "5xA4";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //6xA4
                                    if (
                                        ((sheetHeightMM >= 296.00 && sheetHeightMM <= 298.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1259.00 && sheetWidthMM <= 1261.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 296.00 && sheetWidthMM <= 298.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1259.00 && sheetHeightMM <= 1261.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "6xA4";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //3xA3
                                    if (
                                        ((sheetHeightMM >= 419.00 && sheetHeightMM <= 421.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 890.00 && sheetWidthMM <= 892.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 419.00 && sheetWidthMM <= 421.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 890.00 && sheetHeightMM <= 892.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "3xA3";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //4xA3
                                    if (
                                        ((sheetHeightMM >= 419.00 && sheetHeightMM <= 421.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1187.00 && sheetWidthMM <= 1189.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 419.00 && sheetWidthMM <= 421.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1187.00 && sheetHeightMM <= 1189.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "4xA3";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }

                                    //3xA2
                                    if (
                                        ((sheetHeightMM >= 593.00 && sheetHeightMM <= 595.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1259.00 && sheetWidthMM <= 1261.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 593.00 && sheetWidthMM <= 595.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1259.00 && sheetHeightMM <= 1261.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "3xA2";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //4xA2
                                    if (
                                        ((sheetHeightMM >= 593.00 && sheetHeightMM <= 595.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1679.00 && sheetWidthMM <= 1681.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 593.00 && sheetWidthMM <= 595.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1679.00 && sheetHeightMM <= 1681.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "4xA2";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //3xA1
                                    if (
                                        ((sheetHeightMM >= 840.00 && sheetHeightMM <= 842.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 1781.00 && sheetWidthMM <= 1783.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 840.00 && sheetWidthMM <= 842.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 1781.00 && sheetHeightMM <= 1783.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "3xA1";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //4xA1
                                    if (
                                        ((sheetHeightMM >= 840.00 && sheetHeightMM <= 842.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 2375.00 && sheetWidthMM <= 2377.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 840.00 && sheetWidthMM <= 842.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 2375.00 && sheetHeightMM <= 2377.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "4xA1";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    //3xA0
                                    if (
                                        ((sheetHeightMM >= 1188.00 && sheetHeightMM <= 1190.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 2522.00 && sheetWidthMM <= 2524.00))       // Если ширина приблизительно равна 297
                                        ||                                                        // или
                                        ((sheetWidthMM >= 1188.00 && sheetWidthMM <= 1190.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 2522.00 && sheetHeightMM <= 2524.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "3xA0";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }

                                    /*
                                    //4xA0
                                    if (
                                        ((sheetHeightMM >= 1188.00 && sheetHeightMM <= 1190.00) &&  // Если высота приблизительно равна 420
                                        (sheetWidthMM >= 3363.00 && sheetWidthMM <= 3365.00))       // Если ширина приблизительно равна 297
                                        ||                                                          // или
                                        ((sheetWidthMM >= 1188.00 && sheetWidthMM <= 1190.00) &&    // Если ширина приблизительно равна 420
                                        (sheetHeightMM >= 3363.00 && sheetHeightMM <= 3365.00))     // Если высота приблизительно равна 297
                                        )
                                    {
                                        currentPaperSize = "4xA0";
                                        if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                                        {
                                            errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                                            currentPaperSize = "";
                                        }
                                    }
                                    */

                                    // Если формат подобран верно, печатаем, если нет, выдаем ошибку
                                    if (currentPaperSize == "") 
                                    {
                                        TaskDialog.Show("Ошибка", errorMessage + sheet.sheet.SheetNumber + " - " + sheet.sheet.Name);
                                    }
                                    else
                                    {
                                        PaperSizeSet paperSizeSet = printManager.PaperSizes;
                                        foreach (Autodesk.Revit.DB.PaperSize paperSize in paperSizeSet)
                                        {
                                            if (paperSize.Name.ToString() == currentPaperSize)
                                            {
                                                printParam.PaperSize = paperSize;
                                            }
                                        }

                                        string buildingName = building.Text;
                                        if (buildingName == "<ПРИМЕЧАНИЕ НЕ ЗАДАНО>")
                                        {
                                            buildingName = "";
                                        }
                                        else
                                        {
                                            buildingName += "_";
                                        }

                                        string tomeName = tome.Text;
                                        if (tomeName == "<РАЗДЕЛ НЕ ЗАДАН>")
                                        {
                                            tomeName = "";
                                        }
                                        else
                                        {
                                            tomeName += "_";
                                        }

                                        string filename = buildingName + tomeName + sheet.sheet.SheetNumber.ToString() + "_" + sheet.sheet.Name.ToString();
                                        filename = Regex.Replace(filename, @"\p{C}+", string.Empty); ;
                                        string filePath = SavePath + "\\" + filename + ".pdf";
                                        printManager.PrintSetup.Save();

                                        FindAndFillSaveAsWindow(saveAsDialogCaption, filePath);
                                        printManager.SubmitPrint(sheet.sheet);

                                        //printManager.PrintSetup.Delete();                                        
                                    } // Конец else от проверки формата
                                    
                                        printManager.PrintSetup.Delete();
                                } // Конец else от проверки определения размера листа
                            } // Конец проверки sheetInstance
                        } // Конец проверки галочки напротив листа
                    } // Конец перебора листов
                } // Конец перебора томов
            } // Конец перебора домов
            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);
 
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
       
        public static async Task FindAndFillSaveAsWindow(string windowTitle, string savePath)
        {
            const int WM_SETTEXT = 0x000C;
            const int WM_SETFOCUS = 0x0007;
            const int WM_PASTE = 0x0302;
            const int EM_REPLACESEL = 0x00C2;
            const int BM_CLICK = 0x00F5;
            // Ждем, когда появится окно
            while (FindWindowByCaption(IntPtr.Zero, windowTitle) == IntPtr.Zero)
            {
                await DelayWork(200);
            }
            await DelayWork(200);
            if (File.Exists(savePath))
            {
                File.Delete(savePath);
            }
                // Находим текстбокс
                IntPtr window = FindWindowByCaption(IntPtr.Zero, windowTitle);
            IntPtr dUiView = FindWindowEx(window, IntPtr.Zero, "DUIViewWndClassName", "");
            IntPtr directUiHwnd = FindWindowEx(dUiView, IntPtr.Zero, "DirectUIHWND", "");            
            IntPtr floatNotifySink = FindWindowEx(directUiHwnd, IntPtr.Zero, "FloatNotifySink", string.Empty);
            IntPtr comboBox = FindWindowEx(floatNotifySink, IntPtr.Zero, "ComboBox", "");
            IntPtr textBox= FindWindowEx(comboBox, IntPtr.Zero, "Edit", "");

            // Вставляем в текстбокс путь и имя файла
            SendMessage(textBox, WM_SETFOCUS, IntPtr.Zero, "");
            System.Windows.Forms.Clipboard.SetText(savePath);
            SendMessage(textBox, WM_PASTE, 0, IntPtr.Zero);     

            // Нажимаем на кнопку
            IntPtr button = FindWindowEx(window, IntPtr.Zero, "Button", "Со&хранить");            
            SendMessage(button, BM_CLICK, IntPtr.Zero, null);

        }   
        private static async Task DelayWork(int i)
        {
            await Task.Delay(i);
        }

        public bool CheckPaperSize(string printerName, string paperName)
        {
            bool flag = false;
            PrintDocument pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;

            foreach (System.Drawing.Printing.PaperSize paperSize in pd.PrinterSettings.PaperSizes)
            {
                if (paperName == paperSize.PaperName)
                {
                    flag = true;
                }
            }
            return flag;
        }
    }
}
