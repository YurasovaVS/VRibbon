using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using System.Windows.Automation;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UIFramework;
using Xceed.Wpf.AvalonDock;
using Xceed.Wpf.AvalonDock.Controls;
using Xceed.Wpf.AvalonDock.Layout;
using System.Windows.Interop;
using SKRibbon;

namespace ColorizeTabs
{
    [Transaction(TransactionMode.Manual)]
    public class ColorizeTabs : IExternalCommand
    {
        // Глобальные переменные
        public static Dictionary<long, Brush> DocumentBrushes;

        public static string[] hexes;
        public static List<SolidColorBrush> DocumentBrushThemeColor = new List<SolidColorBrush>();
        // Здесь не нужен лист!!!!
        public static List<SolidColorBrush> DocumentBrushThemeWhite = new List<SolidColorBrush>()
        {
            Brushes.Gainsboro
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            SKRibbon.Properties.appSettings.Default.tabColorFlag = !SKRibbon.Properties.appSettings.Default.tabColorFlag;
            SKRibbon.Properties.appSettings.Default.Save();
            RunCommand(commandData.Application, SKRibbon.Properties.appSettings.Default.tabColorFlag);
            return Result.Succeeded;
        }


        // Доп. методы
        public static Visual GetWindowRoot(UIApplication uiapp)
        {
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = uiapp.MainWindowHandle;
            }
            catch
            {
            }
            return (hwnd != IntPtr.Zero) ? HwndSource.FromHwnd(hwnd).RootVisual : (Visual)null;
        }
        public static LayoutDocumentPaneGroupControl GetDocumentTabGroup(UIApplication uiapp)
        {
            Visual windowRoot = GetWindowRoot(uiapp);
            if (windowRoot == null)
                return (LayoutDocumentPaneGroupControl)null;
            LayoutDocumentPaneGroupControl firstChild = MainWindow.FindFirstChild<LayoutDocumentPaneGroupControl>((DependencyObject)windowRoot);
            // Следующая строчка нужна, чтобы вкладки не перекрашивались при открытии новых документов в другие цвета.
            // "Почему так?" - спросите вы.
            // "А фиг знает," - отвечу я.
            MainWindow.FindFirstChild<DockingManager>((DependencyObject)windowRoot);
            return firstChild;
        }

        public static IEnumerable<LayoutDocumentPaneControl> GetDocumentPanes(LayoutDocumentPaneGroupControl docTabGroup)
        {
            return (docTabGroup != null) ? docTabGroup.FindVisualChildren<LayoutDocumentPaneControl>() : (IEnumerable<LayoutDocumentPaneControl>)new List<LayoutDocumentPaneControl>();
        }

        public static IEnumerable<TabItem> GetDocumentTabs(LayoutDocumentPaneControl docPane)
        {
            return (docPane != null) ? docPane.FindVisualChildren<TabItem>() : (IEnumerable<TabItem>)new List<TabItem>();
        }

        public static long GetAPIDocumentId(Document doc)
        {
            object obj = doc.GetType().GetMethod("getMFCDoc", BindingFlags.Instance | BindingFlags.NonPublic).Invoke((object)doc, new object[0]);
            return ((IntPtr)obj.GetType().GetMethod("GetPointerValue", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(obj, new object[0])).ToInt64();
        }

        public static long GetTabDocumentId(TabItem tab) => ((MFCMDIFrameHost)((ContentControl)((LayoutContent)tab.Content).Content).Content).document.ToInt64();

        public static void RunCommand (UIApplication uiApp, bool flag)
        {
            DocumentBrushThemeColor = GetSavedColors();
            IEnumerable<LayoutDocumentPaneControl> documentPanes = GetDocumentPanes(GetDocumentTabGroup(uiApp));
            DocumentBrushes = new Dictionary<long, Brush>();

            foreach (LayoutDocumentPaneControl docPane in documentPanes)
            {
                Dictionary<long, Brush> dictionary = new Dictionary<long, Brush>();
                IEnumerable<TabItem> documentTabs = GetDocumentTabs(docPane);

                foreach (TabItem tabItem in documentTabs)
                {
                    tabItem.BorderBrush = (Brush)Brushes.White;
                    tabItem.BorderThickness = new Thickness();
                }
                foreach (Document document in uiApp.Application.Documents)
                {
                    if (!document.IsLinked)
                    {
                        long apiDocumentId = GetAPIDocumentId(document);
                        Brush brush1 = (Brush)null;
                        if (DocumentBrushes.ContainsKey(apiDocumentId))
                        {
                            brush1 = DocumentBrushes[apiDocumentId];
                        }
                        else
                        {
                            // if else нас уже не устраивает??? ----------------------------------------------------!!!
                            if (!flag)
                            {
                                // Здесь можно сделать одно присвоение!!!------------------------------------------!!!!
                                foreach (Brush brush2 in DocumentBrushThemeWhite)
                                    brush1 = brush2;
                            }
                            if (flag)
                            {
                                foreach (Brush brush3 in DocumentBrushThemeColor)
                                {
                                    if (!DocumentBrushes.ContainsValue(brush3))
                                    {
                                        brush1 = brush3;
                                        break;
                                    }
                                }
                            }
                            DocumentBrushes[apiDocumentId] = brush1;
                        }
                        if (brush1 != null)
                        {
                            dictionary[apiDocumentId] = brush1;
                            foreach (TabItem tab in documentTabs)
                            {
                                if (GetTabDocumentId(tab) == apiDocumentId)
                                {
                                    if (!flag)
                                    {
                                        if (tab.IsSelected)
                                        {
                                            tab.BorderBrush = (Brush)Brushes.White;
                                            tab.Background = (Brush)Brushes.White;
                                            tab.OpacityMask = (Brush)Brushes.White;
                                        }
                                        else
                                        {
                                            tab.BorderBrush = brush1;
                                            tab.Background = brush1;
                                            tab.OpacityMask = brush1;
                                        }
                                    }
                                    else
                                    {
                                        tab.BorderBrush = brush1;
                                        tab.Background = brush1;
                                        tab.OpacityMask = brush1;
                                        if (document.IsFamilyDocument)
                                            tab.BorderThickness = new Thickness(1.0);
                                        else
                                            tab.BorderThickness = new Thickness(0.0, 1.0, 0.0, 0.0);
                                    }
                                }
                            }
                        }
                    }
                }
                DocumentBrushes = dictionary;
            }
        }
        public static SolidColorBrush HexToBrush (string hex)
        {
            BrushConverter converter = new BrushConverter();
            SolidColorBrush brush = (SolidColorBrush)converter.ConvertFromString(hex);
            return brush;
        }

        public static List<SolidColorBrush> GetSavedColors()
        {
            hexes = SKRibbon.Properties.appSettings.Default.tabColors.Split(',');
            List<SolidColorBrush> brushList = new List<SolidColorBrush>();
            foreach (string hex in hexes)
            {
                brushList.Add(HexToBrush(hex));
            }
            return brushList;
        }
    }
}
