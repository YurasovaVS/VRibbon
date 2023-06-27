using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.Attributes;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using System.Windows.Media;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        static void AddRibbonPanel(UIControlledApplication application)
        {
            String tabName = "АМ СК";
            application.CreateRibbonTab(tabName);

            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Штампы");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            
            PushButtonData b13Data = new PushButtonData(
                "cmdFillStamps",
                "Заполнить" + System.Environment.NewLine + "штамп",
                thisAssemblyPath,
                "FillStamps.FillStamps");

            PushButton pb13 = ribbonPanel.AddItem(b13Data) as PushButton;
            pb13.ToolTip = "Заполняет штампы на выбранных листах.";
            BitmapImage pb13Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/stampFill.png"));
            pb13.LargeImage = pb13Image;


            PushButtonData b1Data = new PushButtonData(
                "cmdAddSignatureDWG",
                "Проставить" + System.Environment.NewLine + "подписи",
                thisAssemblyPath,
                "AddSignatures.AddSignaturesDWG");

            PushButton pb1 = ribbonPanel.AddItem(b1Data) as PushButton;
            pb1.ToolTip = "Проставляет на листах подписи, связанные с внешним DWG файлом. Подписи прикрепляются связью.";
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/add.png"));
            pb1.LargeImage = pb1Image;

            PushButtonData b2Data = new PushButtonData(
                "cmdDeleteSignatureDWG",
                "Удалить" + System.Environment.NewLine + "подписи",
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

            // Кнопка ИУЛов
            PushButtonData b9Data = new PushButtonData(
              "cmdCreateInfoLists",
              "Вывести ИУЛы",
              thisAssemblyPath,
              "InfoListMaker.CreateInfoList");

            PushButton pb9 = rpCopyLists.AddItem(b9Data) as PushButton;
            pb8.ToolTip = "Создает ИУЛы для выбранных томов";
            BitmapImage pb9Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/infoLists.png"));
            pb9.LargeImage = pb9Image;

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

            // Кнопка "Что вы там натворили?!
            PushButtonData b10Data = new PushButtonData(
               "cmdWhatDidTheyDo",
               "Что они там" + System.Environment.NewLine + "натворили?!",
               thisAssemblyPath,
               "FilterByPeople.FilterByPeople");

            PushButton pb10 = rpTeam.AddItem(b10Data) as PushButton;
            pb10.ToolTip = "Фильтрует выделенные элементы по создателю, последнему изменившему или заемщику";
            BitmapImage pb10Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/whatDidTheyDo.png"));
            pb10.LargeImage = pb10Image;

            // Добавляем панель "Всякое"
            RibbonPanel rpSettings = application.CreateRibbonPanel(tabName, "Интерфейс");
            // Кнопка Раскрасить
            PushButtonData b11Data = new PushButtonData(
               "cmdColorizeTabs",
               "Раскрасить" + System.Environment.NewLine + "вкладки",
               thisAssemblyPath,
               "ColorizeTabs.ColorizeTabs");

            PushButton pb11 = rpSettings.AddItem(b11Data) as PushButton;
            pb11.ToolTip = "Раскрашивает вкладки";
            BitmapImage pb11Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/colorizeTabs.png"));
            pb11.LargeImage = pb11Image;

            // Кнопка Раскрасить
            PushButtonData b12Data = new PushButtonData(
               "cmdChangeColorSettings",
               "Настройки" + System.Environment.NewLine + "цветов",
               thisAssemblyPath,
               "SKRibbon.ChangeTabColorSettings");

            PushButton pb12 = rpSettings.AddItem(b12Data) as PushButton;
            pb12.ToolTip = "Изменить настройки цветов";
            BitmapImage pb12Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/palette.png"));
            pb12.LargeImage = pb12Image;

        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            AddRibbonPanel(application);
            application.ViewActivating += new EventHandler<ViewActivatingEventArgs>(this.OnViewActivating);

            return Result.Succeeded;
        }

        private void OnViewActivating(object sender, ViewActivatingEventArgs e) => (sender as UIApplication).Idling += new EventHandler<IdlingEventArgs>(this.OnIdling);

        private void OnViewActivated(object sender, ViewActivatedEventArgs e) => (sender as UIApplication).Idling += new EventHandler<IdlingEventArgs>(this.OnIdling);

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            UIApplication uiApp = sender as UIApplication;
            if (uiApp != null)
            {                
               ColorizeTabs.ColorizeTabs.RunCommand(uiApp, Properties.appSettings.Default.tabColorFlag);
            }
            uiApp.Idling -= new EventHandler<IdlingEventArgs>(this.OnIdling);
        }
        public List<SolidColorBrush> GetSavedColors(string[] colorHexes)
        {
            List<SolidColorBrush> brushList = new List<SolidColorBrush>();
            foreach (string hex in colorHexes)
            {
                brushList.Add(HexToBrush(hex));
            }
            return brushList;
        }
        
        public SolidColorBrush HexToBrush (string hex)
        {
            BrushConverter converter = new BrushConverter();
            SolidColorBrush brush = (SolidColorBrush)converter.ConvertFromString(hex);
            return brush;
        }
    }
}
