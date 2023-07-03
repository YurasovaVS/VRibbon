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
using SKRibbon;
using MJMCustomPrintForm;

namespace BatchPrinting
{

    public partial class BatchPrintForm : System.Windows.Forms.Form
    {
        Document Doc;
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        string SavePath;

        List<SheetSizes> SHEET_SIZES = new List<SheetSizes>() {
            new SheetSizes(297.00, 210.00, "A4"),
            new SheetSizes(420.00, 297.00, "A3"),
            new SheetSizes(594.00, 420.00, "A2"),
            new SheetSizes(841.00, 594.00, "A1"),
            new SheetSizes(1189.00, 841.00, "A0"),
            new SheetSizes(297.00, 630.00, "3xA4"),
            new SheetSizes(297.00, 840.00, "4xA4"),
            new SheetSizes(297.00, 1050.00, "5xA4"),
            new SheetSizes(297.00, 1260.00, "6xA4"),
            new SheetSizes(420.00, 891.00, "3xA3"),
            new SheetSizes(420.00, 1189.00, "4xA3"),
            new SheetSizes(594.00, 1260.00, "2xA2"),
            new SheetSizes(594.00, 1680.00, "4xA2"),
            new SheetSizes(840.00, 1782.00, "3xA1"),
            new SheetSizes(840.00, 2376.00, "4xA1"),
            new SheetSizes(1189.00, 2523.00, "3xA0"),
        };

        public BatchPrintForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            SavePath = Path.GetDirectoryName(doc.PathName);
            if (SavePath == "")
            {
                SavePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            SavePath = Path.Combine(SavePath, "pdf");
            if (!Directory.Exists(SavePath))
            {
                Directory.CreateDirectory(SavePath);
            }
            if (SKRibbon.Properties.appSettings.Default.printFolder.Length == 0)
            {
                SKRibbon.Properties.appSettings.Default.printFolder = SavePath;
                SKRibbon.Properties.appSettings.Default.Save();
            }
            else
            {
                SavePath = SKRibbon.Properties.appSettings.Default.printFolder;
            }
            this.AutoScroll = true;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Создаем дерево листов проекта
            buildingsDict = SKRibbon.FormUtils.CollectSheetDictionary(Doc, true);
            TreeView tree = SKRibbon.FormUtils.CreateSheetTreeView(buildingsDict);

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

            // Добавляем кнопку пути
            Button pathButton = new Button();
            pathButton.Text = "Выбрать другую папку";
            pathButton.Size = new Size(100, 30);
            pathButton.Margin = new Padding(0, 0, 0, 30);
            pathButton.Click += ChooseFolder;
            pathButton.Parent = optionsWrapper;
            optionsWrapper.Controls.Add(pathButton);
            pathButton.Anchor = AnchorStyles.Left;

            // Добавляем кнопку запуска программы
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

        private class SheetSizes
        {
            public double Height;
            public double Width;
            public string Name;
            public SheetSizes(double height, double width, string name)
            {
                Height = height;
                Width = width;
                Name = name;
            }
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

        public void ChooseFolder(object sender, EventArgs e)
        {
            FlowLayoutPanel wrapper = (FlowLayoutPanel)formWrapper.Controls[1];
            Label displayPath = (Label)wrapper.Controls[2];

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                displayPath.Text = dialog.SelectedPath;
                SavePath = dialog.SelectedPath;
            }
        }

        // Обработка дерева и вывод листов
        private void PrintSheets(object sender, EventArgs e)
        {
            TreeView tree = (TreeView)formWrapper.Controls[0];
            FlowLayoutPanel optionsWrapper = (FlowLayoutPanel)formWrapper.Controls[1];
            System.Windows.Forms.ComboBox printersCB = (System.Windows.Forms.ComboBox)optionsWrapper.Controls[1];
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

            StringBuilder sb = new StringBuilder();

            foreach (TreeNode building in tree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (SKRibbon.FormUtils.SheetTreeNode sheet in tome.Nodes)
                    {
                        if (!sheet.Checked)
                        {
                            continue;
                        }
                        FamilyInstance sheetInstance = new FilteredElementCollector(Doc, sheet.sheet.Id).
                                                        OfClass(typeof(FamilyInstance)).
                                                        OfCategory(BuiltInCategory.OST_TitleBlocks).
                                                        FirstOrDefault() as FamilyInstance;
                        if (sheetInstance == null)
                        {
                            continue;
                        }
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
                            sb.AppendLine("Не получилось определить размеры листа " + sheet.sheet.SheetNumber + " - " + sheet.sheet.Name);
                            continue;
                        }

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
                        printParam.PageOrientation = (sheetHeightMM > sheetWidthMM) ? PageOrientationType.Portrait : PageOrientationType.Landscape;

                        string errorMessage = "Нестандартный размер листа ";
                        string currentPaperSize = "";

                        //Ищем, подходит ли размер под стандартные
                        SHEET_SIZES.ForEach(size =>
                            SetPaperSize(printersCB, sheetHeightMM, sheetWidthMM, size.Height, size.Width, size.Name, ref errorMessage, ref currentPaperSize)
                        );

                        // Если формат подобран верно, печатаем, если нет, выдаем ошибку
                        if (currentPaperSize == "")
                        {
                            sb.AppendLine(errorMessage + sheet.sheet.SheetNumber + " - " + sheet.sheet.Name);
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
                            buildingName = (buildingName == "<ПРИМЕЧАНИЕ НЕ ЗАДАНО>") ? "" : buildingName + "_";

                            string tomeName = tome.Text;
                            tomeName = (tomeName == "<РАЗДЕЛ НЕ ЗАДАН>") ? "" : tomeName + "_";

                            string filename = buildingName + tomeName + sheet.sheet.SheetNumber.ToString() + "_" + sheet.sheet.Name.ToString();
                            filename = Regex.Replace(filename, @"\p{C}+", string.Empty); ;
                            filename = Regex.Replace(filename, @"[\~#%&*{}/:<>?|"",;']", string.Empty);
                            string filePath = SavePath + "\\" + filename + ".pdf";
                            printManager.PrintSetup.Save();

                            FindAndFillSaveAsWindow(saveAsDialogCaption, filePath);
                            printManager.SubmitPrint(sheet.sheet);
                        } // Конец else от проверки формата

                        printManager.PrintSetup.Delete();
                    } // Конец перебора листов                    
                } // Конец перебора томов
            } // Конец перебора домов

            if (sb.Length == 0)
            {
                sb.AppendLine("Ошибок не обнаружено");
            }

            TaskDialog.Show("Ошибки", sb.ToString());

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
            await DelayWork(300);
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
            IntPtr textBox = FindWindowEx(comboBox, IntPtr.Zero, "Edit", "");

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

        public void SetPaperSize(System.Windows.Forms.ComboBox printersCB, double sheetHeightMM, double sheetWidthMM, double w, double h, string name, ref string errorMessage, ref string currentPaperSize)
        {
            if (
                ((sheetHeightMM >= w - 1 && sheetHeightMM <= w + 1) &&  // Если высота приблизительно равна 
                (sheetWidthMM >= h - 1 && sheetWidthMM <= h + 1))       // Если ширина приблизительно равна 
                ||                                                        // или
                ((sheetWidthMM >= w - 1 && sheetWidthMM <= w + 1) &&    // Если ширина приблизительно равна 
                (sheetHeightMM >= h - 1 && sheetHeightMM <= h + 1))     // Если высота приблизительно равна 
               )
            {
                currentPaperSize = name;
                if (!CheckPaperSize(printersCB.SelectedItem.ToString(), currentPaperSize))
                {
                    errorMessage = "Принтер не поддерживает формат " + currentPaperSize + " | ";
                    currentPaperSize = "";
                }
            }
        }

    }
}