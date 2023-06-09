using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKRibbon.Forms
{
    public partial class DeleteSigForm : System.Windows.Forms.Form
    {
        Document Doc;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        public DeleteSigForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;

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
            header.Text = "Выберите листы, с которых хотите удалить подписи:";

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

            //Добавляем кнопку ОК
            Button button = new Button();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Width = 200;

            button.Text = "Убрать подписи";
            button.Click += RemoveSignatures;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;
        }

        public void RemoveSignatures (object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;
            TreeView sheetTree = (TreeView)formWrapper.Controls[1];

            Transaction t = new Transaction(Doc, "Убрать подписи");
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
                            ICollection<Element> signatures = new FilteredElementCollector(Doc, sheet.Id).
                                                            OfClass(typeof(ImportInstance)).
                                                            ToElements();
                            foreach (ImportInstance signature in signatures)
                            {
                                string name = signature.LookupParameter("Имя").AsString();
                                string pattern = "подпись_*";
                                if (Regex.Match(name, pattern).Success)
                                {
                                    Doc.Delete(signature.Id);
                                }
                            }
                        }
                    }
                }
            }

            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        public void CheckAllChildNodes(TreeNode node, bool nodeChecker)
        {
            foreach (TreeNode childNode in node.Nodes)
            {
                childNode.Checked = nodeChecker;
                if (childNode.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(childNode, nodeChecker);
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
    public class CheckedSheetList : CheckedListBox
    {
        public List<Autodesk.Revit.DB.ViewSheet> sheetCollection;
        public CheckedSheetList()
        {
            sheetCollection = new List<Autodesk.Revit.DB.ViewSheet>();
        }
    }
}
