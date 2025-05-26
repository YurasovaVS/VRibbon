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
        UIApplication uiApp_cached = null;
        static void AddRibbonPanel(UIControlledApplication application)
        {
            string version = "v2.0";

            String tabName = "Витрувий," + System.Environment.NewLine + "помоги!";
            application.CreateRibbonTab(tabName);
            
            // Создаем панельки на вкладке
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Штампы");
            RibbonPanel rpAreas = application.CreateRibbonPanel(tabName, "Скорректировать площади");
            RibbonPanel rpCopyLists = application.CreateRibbonPanel(tabName, "Листы");
            RibbonPanel rpModellingTools = application.CreateRibbonPanel(tabName, "Модель");
            RibbonPanel rpEngineers = application.CreateRibbonPanel(tabName, "Смежники");
            RibbonPanel rpTeam = application.CreateRibbonPanel(tabName, "Злой начальник");
            RibbonPanel rpSettings = application.CreateRibbonPanel(tabName, "Интерфейс");
            RibbonPanel rpInfo = application.CreateRibbonPanel(tabName, "Информация");


            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            // ------------------------------------------------------
            
            IList<RibbonItem> stackedGroup1 = ribbonPanel.AddStackedItems(
                AddStackedButton("cmdFillStamps",
                            "Заполнить штамп",
                            thisAssemblyPath,
                            "FillStamps.FillStamps",
                            "stampFill_16.png",
                            "stampFill.png",
                            "Заполняет штампы на выбранных листах."
                            ),
                AddStackedButton("cmdAddSignatureDWG",
                            "Проставить подписи",
                            thisAssemblyPath,
                            "AddSignatures.AddSignaturesDWG",
                            "add_16.png",
                            "add.png",
                            "Проставляет на листах подписи, связанные с внешним DWG файлом. Подписи прикрепляются связью."
                            ),
                AddStackedButton("cmdDeleteSignatureDWG",
                            "Удалить подписи",
                            thisAssemblyPath,
                            "DeleteSignatures.DeleteSignaturesDWG",
                            "delete_16.png",
                            "delete.png",
                            "Удаляет на выбранных листах подписи, связанные с внешним DWG файлом."
                            )
                );           

          
            // ------------------------------------------------------
            //Кнопка рассчета новых площадей
            AddPushPutton(rpAreas,
                            "cmdAdjustAreas",
                            "Исправить" + System.Environment.NewLine + " площади ",
                            thisAssemblyPath,
                            "FakeArea.AdjustExistingAreas",
                            "fakeArea.png",
                            "Корректирует площади помещений выбранной категории (по функции), приводя их общую площадь к искомому значению. Новая площадь прописывается в параметр \"Комментарий\"",
                            true
                            );
            //Кнопка замены марок на определенных видах
            AddPushPutton(rpAreas,
                            "cmdReplaceTagFamilies",
                            "Заменить" + System.Environment.NewLine + " марки ",
                            thisAssemblyPath,
                            "ReplaceAreaTags.ReplaceTags",
                            "replaceTags.png",
                            "Заменяет марки определенного типа на нужных листах",
                            true
                            );
            // ------------------------------------------------------

            IList<RibbonItem> stackedGroup2 = rpCopyLists.AddStackedItems(
                AddStackedButton("cmdCopyLists",
                            "Дублировать в проекте ",
                            thisAssemblyPath,
                            "CopyListsTree.CopyLists",
                            "copyLists_16.png",
                            "copyLists.png",
                            "По выбору пользователя дублирует существующие в проекте листы"
                            ),
                AddStackedButton("cmdPrintToPdf",
                            "Вывести в PDF",
                            thisAssemblyPath,
                            "BatchPrinting.BatchPrintSheets",
                            "pdf_16.png",
                            "pdf.png",
                            "Вывод листов в PDF"
                            ),
                AddStackedButton("cmdExportSheets",
                            "Вывести в DWG",
                            thisAssemblyPath,
                            "SKRibbon.BatchDwgExport",
                            "dwgExport_16.png",
                            "dwgExport.png",
                            "Экспортирует выбранные листы в DWG"
                            )
                );

            // Кнопка Ренумерации
            AddPushPutton(rpCopyLists,
                            "cmdRenumSheets",
                            "Ренумерация",
                            thisAssemblyPath,
                            "SheetRenamer.RenameSheets",
                            "renumber.png",
                            "Меняет нумерацию листов в выбранном томе, начиная с выбранного листа",
                            true
                            );

            // Кнопка ИУЛов
            AddPushPutton(rpCopyLists,
                            "cmdCreateInfoLists",
                            "Вывести ИУЛы",
                            thisAssemblyPath,
                            "InfoListMaker.CreateInfoList",
                            "infoLists.png",
                            "Создает ИУЛы для выбранных томов",
                            true
                            );
            // ------------------------------------------------------
            // Кнопка Исправить отзеркаленные двери
            AddPushPutton(rpModellingTools,
                            "cmdFixMirroredDoors",
                            "Исправить двери",
                            thisAssemblyPath,
                            "SKRibbon.FixMirroredDoors",
                            "fixMirroredDoors.png",
                            "Исправляет отзеркаленные двери",
                            true
                            );

            // Кнопка расставить полы
            AddPushPutton(rpModellingTools,
                            "cmdPlaceFloors",
                            "Создать полы",
                            thisAssemblyPath,
                            "SKRibbon.PlaceFloors",
                            "floor.png",
                            "Создать полы в комнатах",
                            true
                            );

            // Кнопка нумерации помещений
            AddPushPutton(rpModellingTools,
                            "cmdNumerateRooms",
                            "Пронумеровать" + System.Environment.NewLine + "помещения",
                            thisAssemblyPath,
                            "SKRibbon.NumerateRooms",
                            "numerateRooms.png",
                            "Пронумеровать комнаты по методике метрополитена",
                            true
                            );

            // Кнопка распределения по рабочим наборам
            AddPushPutton(rpModellingTools,
                            "cmdFixWorkGroups",
                            "Исправить" + System.Environment.NewLine + "наборы",
                            thisAssemblyPath,
                            "SKRibbon.FixWorkGroups",
                            "sortIntoWorkgroups_32.png",
                            "Проверить и исправить соответствие элементов рабочим наборам",
                            true
                            );

            // ------------------------------------------------------
            // Кнопка исправления файлов IFC
            AddPushPutton(rpEngineers,
                            "cmdFixWorkGroups",
                            "Исправить IFC",
                            thisAssemblyPath,
                            "SKRibbon.FixIFCCoordinates",
                            "fixIFCCoord_32.png",
                            "Исправить координаты в IFC файле",
                            true
                            );


            // ------------------------------------------------------
            // Кнопка "Кто это сделал?!
            AddPushPutton(rpTeam,
                            "cmdWhoDidThat",
                            "Кто это" + System.Environment.NewLine + "сделал?!",
                            thisAssemblyPath,
                            "WhoDidThat.WhoDidThat",
                            "whoDidThat.png",
                            "Показывает, кто сделал, и кто последним изменил выбранные элементы",
                            true
                            );
            // Кнопка "Что вы там натворили?!"
            AddPushPutton(rpTeam,
                            "cmdWhatDidTheyDo",
                            "Что они там" + System.Environment.NewLine + "натворили?!",
                            thisAssemblyPath,
                            "FilterByPeople.FilterByPeople",
                            "whatDidTheyDo.png",
                            "Фильтрует выделенные элементы по создателю, последнему изменившему или заемщику",
                            true
                            );
            // ------------------------------------------------------
            // Кнопка "Раскрасить вкладки"
            AddPushPutton(rpSettings,
                            "cmdColorizeTabs",
                            "Раскрасить" + System.Environment.NewLine + "вкладки",
                            thisAssemblyPath,
                            "ColorizeTabs.ColorizeTabs",
                            "colorizeTabs.png",
                            "Раскрашивает вкладки",
                            true
                            );
            // Кнопка "Настроить цвета"
            AddPushPutton(rpSettings,
                            "cmdChangeColorSettings",
                            "Настройки" + System.Environment.NewLine + "цветов",
                            thisAssemblyPath,
                            "SKRibbon.ChangeTabColorSettings",
                            "palette.png",
                            "Изменить настройки цветов",
                            true
                            );

            //------------------------------------------------------
            // Кнопка "О программе"
            AddPushPutton(rpInfo,
                            "cmdInfo",
                            "О программе",
                            thisAssemblyPath,
                            "SKRibbon.Functions.Info",
                            "info.png",
                            "Важная информация",
                            true
                            );
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
            // Кручу-верчу на DocumentChanged подписаться хочу
            /*
            if (uiApp_cached == null)
            {
                uiApp_cached = uiApp;
                uiApp.Application.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(this.OnDocumentChanged);
            }*/
        }
        /*
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            Document doc = e.GetDocument();
            ICollection addedElements e.GetAddedElementIds();
        }
          */      

        public List<SolidColorBrush> GetSavedColors(string[] colorHexes)
        {
            List<SolidColorBrush> brushList = new List<SolidColorBrush>();
            foreach (string hex in colorHexes)
            {
                brushList.Add(HexToBrush(hex));
            }
            return brushList;
        }
        
        public SolidColorBrush HexToBrush(string hex)
        {
            BrushConverter converter = new BrushConverter();
            SolidColorBrush brush = (SolidColorBrush)converter.ConvertFromString(hex);
            return brush;
        }

        private static void AddPushPutton(RibbonPanel ribbonPanel, string cmdName, string cmdTitle, string thisAssemblyPath, string moduleName, string imgName, string tooltip, bool enabled)
        {
            PushButtonData bData = new PushButtonData(
                cmdName,
                cmdTitle,
                thisAssemblyPath,
                moduleName);

            PushButton pb = ribbonPanel.AddItem(bData) as PushButton;
            pb.ToolTip = tooltip;
            string imgPath = "pack://application:,,,/SKRibbon;component/Resources/" + imgName;
            BitmapImage pbImage = new BitmapImage(new Uri(imgPath));
            pb.LargeImage = pbImage;
            pb.Enabled = enabled;
        }

        private static PushButtonData AddStackedButton (string cmdName, string cmdTitle, string thisAssemblyPath, string moduleName, string smallImg, string largeImg, string tooltip)
        {
            PushButtonData pbd = new PushButtonData(
                                        cmdName,
                                        cmdTitle,
                                        thisAssemblyPath,
                                        moduleName);
            pbd.ToolTip = tooltip;  // Can be changed to a more descriptive text.
            string smallImgPath = "pack://application:,,,/SKRibbon;component/Resources/" + smallImg;
            string largeImgPath = "pack://application:,,,/SKRibbon;component/Resources/" + largeImg;
            pbd.Image = new BitmapImage(new Uri(smallImgPath));
            pbd.LargeImage = new BitmapImage(new Uri(largeImgPath));
            return pbd;
        }
    }
}
