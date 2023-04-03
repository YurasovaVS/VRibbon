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
            pb1.ToolTip = "Проставляет на листах подписи, связанные с внешним DWG файлом. Файлы подписей лежат в папке: Z:\\13_Пользователи\\_ПодписиСК";
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/add.png"));
            pb1.LargeImage = pb1Image;

            PushButtonData b2Data = new PushButtonData(
                "cmdDeleteSignatureDWG",
                "Удалить" + System.Environment.NewLine + "  DWG  ",
                thisAssemblyPath,
                "DeleteSignatures.DeleteSignaturesDWG");

            PushButton pb2 = ribbonPanel.AddItem(b2Data) as PushButton;
            pb2.ToolTip = "Удаляет на листах подписи, связанные с внешним DWG файлом.";
            BitmapImage pb2Image = new BitmapImage(new Uri("pack://application:,,,/SKRibbon;component/Resources/delete.png"));
            pb2.LargeImage = pb2Image;

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
