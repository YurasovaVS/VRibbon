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
            
            // Создаем панельки на вкладке
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Штампы");
            RibbonPanel rpAreas = application.CreateRibbonPanel(tabName, "Скорректировать площади");
            RibbonPanel rpCopyLists = application.CreateRibbonPanel(tabName, "Листы");
            RibbonPanel rpTeam = application.CreateRibbonPanel(tabName, "Злой начальник");
            RibbonPanel rpSettings = application.CreateRibbonPanel(tabName, "Интерфейс");

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

            /*
            // Кнопка заполнения штампов
            AddPushPutton(ribbonPanel, 
                            "cmdFillStamps", 
                            "Заполнить" + System.Environment.NewLine + "штамп", 
                            thisAssemblyPath, 
                            "FillStamps.FillStamps", 
                            "stampFill.png",
                            "Заполняет штампы на выбранных листах."
                            );
            // Кнопка проставления подписей
            AddPushPutton(ribbonPanel,
                            "cmdAddSignatureDWG",
                            "Проставить" + System.Environment.NewLine + "подписи",
                            thisAssemblyPath,
                            "AddSignatures.AddSignaturesDWG",
                            "add.png",
                            "Проставляет на листах подписи, связанные с внешним DWG файлом. Подписи прикрепляются связью."
                            );
            // Кнопка удаления подписей
            AddPushPutton(ribbonPanel,
                            "cmdDeleteSignatureDWG",
                            "Удалить" + System.Environment.NewLine + "подписи",
                            thisAssemblyPath,
                            "DeleteSignatures.DeleteSignaturesDWG",
                            "delete.png",
                            "Удаляет на выбранных листах подписи, связанные с внешним DWG файлом."
                            );
            */

            // ------------------------------------------------------
            //Кнопка рассчета новых площадей
            AddPushPutton(rpAreas,
                            "cmdAdjustAreas",
                            "Исправить" + System.Environment.NewLine + " площади ",
                            thisAssemblyPath,
                            "FakeArea.AdjustExistingAreas",
                            "fakeArea.png",
                            "Корректирует площади помещений выбранной категории (по функции), приводя их общую площадь к искомому значению. Новая площадь прописывается в параметр \"Комментарий\""
                            );
            //Кнопка замены марок на определенных видах
            AddPushPutton(rpAreas,
                            "cmdReplaceTagFamilies",
                            "Заменить" + System.Environment.NewLine + " марки ",
                            thisAssemblyPath,
                            "ReplaceAreaTags.ReplaceTags",
                            "replaceTags.png",
                            "Заменяет марки определенного типа на нужных листах"
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
                AddStackedButton("cmdRenumSheets",
                            "Ренумерация",
                            thisAssemblyPath,
                            "SheetRenamer.RenameSheets",
                            "renumber_16.png",
                            "renumber.png",
                            "Меняет нумерацию листов в выбранном томе, начиная с выбранного листа"
                            )
                );

            /*
            // Кнопка копирования листов
            AddPushPutton(rpCopyLists,
                            "cmdCopyLists",
                            "Дублировать" + System.Environment.NewLine + " в проекте ",
                            thisAssemblyPath,
                            "CopyListsTree.CopyLists",
                            "copyLists.png",
                            "По выбору пользователя дублирует существующие в проекте листы"
                            );
            // Кнопка печати в PDF
            AddPushPutton(rpCopyLists,
                            "cmdPrintToPdf",
                            "Вывести в PDF",
                            thisAssemblyPath,
                            "BatchPrinting.BatchPrintSheets",
                            "pdf.png",
                            "Вывод листов в PDF"
                            );
            // Кнопка перенумерации листов
            AddPushPutton(rpCopyLists,
                            "cmdRenumSheets",
                            "Ренумерация",
                            thisAssemblyPath,
                            "SheetRenamer.RenameSheets",
                            "renumber.png",
                            "Меняет нумерацию листов в выбранном томе, начиная с выбранного листа"
                            );
            */

            // Кнопка ИУЛов
            AddPushPutton(rpCopyLists,
                            "cmdCreateInfoLists",
                            "Вывести ИУЛы",
                            thisAssemblyPath,
                            "InfoListMaker.CreateInfoList",
                            "infoLists.png",
                            "Создает ИУЛы для выбранных томов"
                            );
            // ------------------------------------------------------
            // Кнопка "Кто это сделал?!
            AddPushPutton(rpTeam,
                            "cmdWhoDidThat",
                            "Кто это" + System.Environment.NewLine + "сделал?!",
                            thisAssemblyPath,
                            "WhoDidThat.WhoDidThat",
                            "whoDidThat.png",
                            "Показывает, кто сделал, и кто последним изменил выбранные элементы"
                            );
            // Кнопка "Что вы там натворили?!"
            AddPushPutton(rpTeam,
                            "cmdWhatDidTheyDo",
                            "Что они там" + System.Environment.NewLine + "натворили?!",
                            thisAssemblyPath,
                            "FilterByPeople.FilterByPeople",
                            "whatDidTheyDo.png",
                            "Фильтрует выделенные элементы по создателю, последнему изменившему или заемщику"
                            );
            // ------------------------------------------------------
            // Кнопка "Раскрасить вкладки"
            AddPushPutton(rpSettings,
                            "cmdColorizeTabs",
                            "Раскрасить" + System.Environment.NewLine + "вкладки",
                            thisAssemblyPath,
                            "ColorizeTabs.ColorizeTabs",
                            "colorizeTabs.png",
                            "Раскрашивает вкладки"
                            );
            // Кнопка "Настроить цвета"
            AddPushPutton(rpSettings,
                            "cmdChangeColorSettings",
                            "Настройки" + System.Environment.NewLine + "цветов",
                            thisAssemblyPath,
                            "SKRibbon.ChangeTabColorSettings",
                            "palette.png",
                            "Изменить настройки цветов"
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
        
        public SolidColorBrush HexToBrush(string hex)
        {
            BrushConverter converter = new BrushConverter();
            SolidColorBrush brush = (SolidColorBrush)converter.ConvertFromString(hex);
            return brush;
        }

        private static void AddPushPutton(RibbonPanel ribbonPanel, string cmdName, string cmdTitle, string thisAssemblyPath, string moduleName, string imgName, string tooltip)
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
