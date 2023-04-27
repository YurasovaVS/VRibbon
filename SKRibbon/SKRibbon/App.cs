using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        static void AddRibbonPanel(UIControlledApplication application)
        {
            String tabName = "АМ СК";
            application.CreateRibbonTab(tabName);

            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Подписи");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData b1Data = new PushButtonData(
                "cmdAddSignatureDWG",
                "Проставить" + System.Environment.NewLine + "  DWG  ",
                thisAssemblyPath,
                "AddSignatures.AddSignaturesDWG");

            PushButton pb1 = ribbonPanel.AddItem(b1Data) as PushButton;
            pb1.ToolTip = "Проставляет на листах подписи, связанные с внешним DWG файлом. Подписи прикрепляются связью.";
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/add.png"));
            pb1.LargeImage = pb1Image;

            PushButtonData b2Data = new PushButtonData(
                "cmdDeleteSignatureDWG",
                "Удалить" + System.Environment.NewLine + "  DWG  ",
                thisAssemblyPath,
                "DeleteSignatures.DeleteSignaturesDWG");

            PushButton pb2 = ribbonPanel.AddItem(b2Data) as PushButton;
            pb2.ToolTip = "Удаляет на выбранных листах подписи, связанные с внешним DWG файлом.";
            BitmapImage pb2Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/delete.png"));
            pb2.LargeImage = pb2Image;
            
            //Добавляем панель с изменением площадей
            RibbonPanel rpAreas = application.CreateRibbonPanel(tabName, "Скорректировать площади");
            //Кнопка рассчета новых площадей
            PushButtonData b3Data = new PushButtonData(
               "cmdAdjustAreas",
               "Исправить" + System.Environment.NewLine + " площади ",
               thisAssemblyPath,
               "FakeArea.AdjustExistingAreas");

            PushButton pb3 = rpAreas.AddItem(b3Data) as PushButton;
            pb3.ToolTip = "Корректирует площади помещений выбранной категории (по функции), приводя их общую площадь к искомому значению. Новая площадь прописывается в параметр \"Комментарий\"";
            BitmapImage pb3Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/fakeArea.png"));
            pb3.LargeImage = pb3Image;
            
            //Кнопка замены марок на определенных видах
            PushButtonData b4Data = new PushButtonData(
               "cmdReplaceTagFamilies",
               "Заменить" + System.Environment.NewLine + " марки ",
               thisAssemblyPath,
               "ReplaceAreaTags.ReplaceTags");

            PushButton pb4 = rpAreas.AddItem(b4Data) as PushButton;
            pb4.ToolTip = "Заменяет марки определенного типа на нужных листах";
            BitmapImage pb4Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/replaceTags.png"));
            pb4.LargeImage = pb4Image;

            // Добавляем панель копирования листов
            RibbonPanel rpCopyLists = application.CreateRibbonPanel(tabName, "Листы");
            // Кнопка копирования листов
            PushButtonData b5Data = new PushButtonData(
               "cmdCopyLists",
               "Дублировать" + System.Environment.NewLine + " в проекте ",
               thisAssemblyPath,
               "CopyListsTree.CopyLists");

            PushButton pb5 = rpCopyLists.AddItem(b5Data) as PushButton;
            pb5.ToolTip = "По выбору пользователя дублирует существующие в проекте листы";
            BitmapImage pb5Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/copyLists.png"));
            pb5.LargeImage = pb5Image;

            // Кнопка печати в PDF
            PushButtonData b7Data = new PushButtonData(
              "cmdPrintToPdf",
              "Вывести в PDF",
              thisAssemblyPath,
              "BatchPrinting.BatchPrintSheets");

            PushButton pb7 = rpCopyLists.AddItem(b7Data) as PushButton;
            pb7.ToolTip = "Вывод листов в PDF";
            BitmapImage pb7Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/pdf.png"));
            pb7.LargeImage = pb7Image;

            // Кнопка перенумерации листов
            PushButtonData b8Data = new PushButtonData(
              "cmdRenumSheets",
              "Ренумерация",
              thisAssemblyPath,
              "SheetRenamer.RenameSheets");

            PushButton pb8 = rpCopyLists.AddItem(b8Data) as PushButton;
            pb8.ToolTip = "Меняет нумерацию листов в выбранном томе, начиная с выбранноо листа";
            BitmapImage pb8Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/renumber.png"));
            pb8.LargeImage = pb8Image;

            // Добавляем панель "Злой начальник"
            RibbonPanel rpTeam = application.CreateRibbonPanel(tabName, "Злой начальник");
            // Кнопка "Кто это сделал?!
            PushButtonData b6Data = new PushButtonData(
               "cmdWhoDidThat",
               "Кто это" + System.Environment.NewLine + "сделал?!",
               thisAssemblyPath,
               "WhoDidThat.WhoDidThat");

            PushButton pb6 = rpTeam.AddItem(b6Data) as PushButton;
            pb6.ToolTip = "Показывает, кто сделал, и кто последним изменил выбранные элементы";
            BitmapImage pb6Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/whoDidThat.png"));
            pb6.LargeImage = pb6Image;


        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            AddRibbonPanel(application);
            return Result.Succeeded;
        }
    }
}
