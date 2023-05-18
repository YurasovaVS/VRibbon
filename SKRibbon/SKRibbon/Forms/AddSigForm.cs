using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Security.AccessControl;
using System.IO;

namespace SKRibbon
{
    public partial class AddSigForm : System.Windows.Forms.Form
    {
        Document Doc;
        string Path;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        public AddSigForm(Document doc, string path)
        {
            InitializeComponent();
            Doc = doc;
            Path = path;
            this.AutoSize = true;
            this.AutoScroll = true;

            //Создаем Wrapper для содержимого формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.BorderStyle = BorderStyle.FixedSingle;
            formWrapper.Padding = new Padding(5, 5, 5, 5);

            //Создаем заголовок
            Label header = new Label();
            header.Parent = formWrapper;
            formWrapper.Controls.Add(header);
            header.Anchor = AnchorStyles.Top;
            header.Size = new Size(500, 30);
            header.Text = "Введите путь к папке, где лежат DWG-подписи:";

            //Создаем текстовое поле для пути
            System.Windows.Forms.TextBox newPath = new System.Windows.Forms.TextBox();
            newPath.Parent = formWrapper;
            formWrapper.Controls.Add(newPath);
            newPath.Size = new Size(300, 30);
            newPath.Text = Path;

            //Создаем заголовок чеклиста
            Label checkHeader = new Label();
            checkHeader.Parent = formWrapper;
            formWrapper.Controls.Add(checkHeader);
            checkHeader.Anchor = AnchorStyles.Top;
            checkHeader.Size = new Size(500, 30);
            checkHeader.Text = "Выберите листы:";

            // Добавляем древо листов
            TreeView sheetTree = new TreeView();
            sheetTree.Parent = formWrapper;
            formWrapper.Controls.Add(sheetTree);
            sheetTree.Anchor = AnchorStyles.Top;
            sheetTree.MinimumSize = new Size(500, 30);
            sheetTree.Height = 200;
            sheetTree.CheckBoxes = true;
            sheetTree.AfterCheck += node_AfterCheck;
            

            //Добавляем виды в чеклист
            IEnumerable<Element> sheets = new FilteredElementCollector(doc).
                                            OfCategory(BuiltInCategory.OST_Sheets).
                                            WhereElementIsNotElementType().
                                            ToElements();

            Dictionary<string, Dictionary<string, List<ViewSheet>>> tempDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
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
                    if (!tempDict.ContainsKey(building))
                    {
                        Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                        tempDict.Add(building, tomesDict);
                    }
                    // Если во вложенном словаре здания нет такого тома, создаем его
                    if (!tempDict[building].ContainsKey(tome))
                    {
                        List<ViewSheet> sheetsList = new List<ViewSheet>();
                        tempDict[building].Add(tome, sheetsList);
                    }
                    // Добавляем лист в нужный том
                    tempDict[building][tome].Add(sheet);
                }
            } // Конец создания словаря листов

            // ОТСОРТИРОВАТЬ ВСЕ СПИСКИ
            foreach (var building in tempDict)
            {
                Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                buildingsDict.Add(building.Key, tomesDict);
                foreach (var tome in building.Value)
                {
                    List<ViewSheet> sheetsList = tome.Value.OrderBy(sheet => sheet.SheetNumber).ToList();
                    buildingsDict[building.Key].Add(tome.Key, sheetsList);
                }
            }

            foreach (var building in buildingsDict)
            {
                sheetTree.Nodes.Add(building.Key);
                foreach (var tome in building.Value)
                {
                    int i = sheetTree.Nodes.Count - 1;
                    sheetTree.Nodes[i].Nodes.Add(tome.Key);
                    foreach (ViewSheet sheet in tome.Value)
                    {
                        int j = sheetTree.Nodes[i].Nodes.Count - 1;
                        SheetNode newNode = new SheetNode();
                        newNode.sheet = sheet;
                        newNode.Text = sheet.SheetNumber + " - " + sheet.Name;
                        sheetTree.Nodes[i].Nodes[j].Nodes.Add(newNode);
                    }
                }
            }

            //Добавляем кнопку
            Button button = new Button();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Width = 200;

            button.Text = "Проставить подписи";
            button.Click += PlaceSignatures;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;
        }

        public void PlaceSignatures (object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;

            System.Windows.Forms.TextBox pathBox = (System.Windows.Forms.TextBox)formWrapper.Controls[1];
            string path = pathBox.Text;
            TreeView sheetTree = (TreeView)formWrapper.Controls[3];
            StringBuilder sb = new StringBuilder();

            Transaction t = new Transaction(Doc, "Вставить подписи");
            t.Start();

            foreach (TreeNode building in sheetTree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (SheetNode sheetNode in tome.Nodes)
                    {
                        if (sheetNode.Checked)
                        {
                            ViewSheet sheet = sheetNode.sheet;
                            DWGImportOptions importOptions = new DWGImportOptions();

                            Double sheetRBpointX = sheet.Outline.Max.U;
                            Double sheetRBpointY = sheet.Outline.Min.V;

                            for (int i = 0; i <= 5; i++)
                            {
                                string paramName = "ADSK_Штамп Строка " + (i + 1).ToString() + " фамилия";
                                string paramValue = sheet.LookupParameter(paramName).AsString();

                                // Если фамилия задана
                                if (paramValue != null)
                                {
                                    if (paramValue.Length != 0)
                                    {
                                        string signaturePath = path + "подпись_" + paramValue + ".dwg";
                                        if (File.Exists(signaturePath))
                                        {
                                            ElementId signatureId;
                                            bool flag = Doc.Link(signaturePath, importOptions, sheet, out signatureId);
                                            Element signature = Doc.GetElement(signatureId);
                                            signature.Pinned = false;
                                            Double x = sheetRBpointX - 0.49;
                                            Double y = 30 - i * 5;
                                            y = UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters);
                                            y = sheetRBpointY + y;
                                            XYZ move = new XYZ(x, y, 0);
                                            signature.Location.Move(move);
                                        }
                                        else
                                        {
                                            sb.AppendLine("Подпись не найдена: " + signaturePath);
                                        }
                                    }
                                }
                            } // Конец перебора строк штампа
                        } // Конец if(sheetNode.Checked)
                    } // Конец foreach (sheet in tome)
                } // Конец foreach (tome in building)
            } // Конец foreach (building in tree)

            if (sb.Length == 0)
            {
                sb.AppendLine("Подписи проставлены. Ошибок нет.");
            }
            TaskDialog.Show("Ошибки", sb.ToString());

            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }          
        
        public void CheckAllChildNodes (TreeNode node, bool nodeChecker)
        {
            foreach (TreeNode childNode in node.Nodes)
            {
                childNode.Checked = nodeChecker;
                if (childNode.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes (childNode, nodeChecker);
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
    }

    //Класс для чеклиста листов
    public class CheckedSheetList : CheckedListBox
    {
        public List<Autodesk.Revit.DB.ViewSheet> sheetCollection;
        public CheckedSheetList()
        {
            sheetCollection = new List<Autodesk.Revit.DB.ViewSheet>();
        }
    }

    // Класс для нодов с листами
    public class SheetNode : TreeNode
    {
        public ViewSheet sheet;
    }
}
