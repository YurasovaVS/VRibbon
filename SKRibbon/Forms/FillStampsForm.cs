using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace FillStamps
{
    [Transaction(TransactionMode.Manual)]
    public partial class FillStampsForm : WinForms.Form
    {
        Document Doc;
        FlowLayoutPanel formWrapper;
        FlowLayoutPanel linesWrapper;
        WinForms.CheckBox checkBox;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();

        public FillStampsForm(Document doc)
        {
            InitializeComponent();

            this.Text = "Заполнить штампы";
            Doc = doc;

            formWrapper = new FlowLayoutPanel();
            formWrapper.FlowDirection = FlowDirection.LeftToRight;
            formWrapper.AutoSize = true;

            linesWrapper = new FlowLayoutPanel();
            linesWrapper.FlowDirection = FlowDirection.TopDown;
            linesWrapper.AutoSize = true;


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
            tree.Anchor = AnchorStyles.Left;

            //-----------------------------------------------------------------
            // Инициализация шаблона штампа

            for (int i = 1; i <= 6; i++)
            {
                FlowLayoutPanel lineWrapper = new FlowLayoutPanel();
                lineWrapper.FlowDirection = FlowDirection.LeftToRight;
                lineWrapper.AutoSize = true;

                Label posLabel = new Label();
                WinForms.TextBox posText = new WinForms.TextBox();
                Label nameLabel = new Label();
                WinForms.TextBox nameText = new WinForms.TextBox();

                posLabel.Size = new System.Drawing.Size(100, 30);
                posText.Size = new System.Drawing.Size(100, 30);
                nameLabel.Size = new System.Drawing.Size(100, 30);
                nameText.Size = new System.Drawing.Size(100, 30);

                posLabel.Anchor = AnchorStyles.Left;
                posText.Anchor = AnchorStyles.Left;
                nameLabel.Anchor = AnchorStyles.Left;
                nameText.Anchor = AnchorStyles.Left;

                posLabel.Text = "Должность " + i.ToString();
                nameLabel.Text = "Фамилия " + i.ToString();

                lineWrapper.Controls.Add(posLabel);
                lineWrapper.Controls.Add(posText);
                lineWrapper.Controls.Add(nameLabel);
                lineWrapper.Controls.Add(nameText);

                posLabel.Parent = lineWrapper;
                posText.Parent = lineWrapper;
                nameLabel.Parent = lineWrapper;
                nameText.Parent = lineWrapper;

                lineWrapper.Parent = linesWrapper;
                linesWrapper.Controls.Add (lineWrapper);
            }

            checkBox = new WinForms.CheckBox();
            checkBox.Size = new System.Drawing.Size (200, 60);
            checkBox.Anchor = AnchorStyles.Top;
            checkBox.Text = "Учитывать (перезаписывать) пустые поля";

            Button button = new Button();
            button.Size = new System.Drawing.Size(150, 100);
            button.Anchor = AnchorStyles.Top;
            button.Text = "Заполнить штампы";
            button.Click += RunProgram;

            checkBox.Parent = linesWrapper;
            linesWrapper.Controls.Add(checkBox);

            button.Parent = linesWrapper;
            linesWrapper.Controls.Add(button);

            tree.Parent = formWrapper;
            formWrapper.Controls.Add(tree);

            linesWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(linesWrapper);

            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);
        }

        private class SheetTreeNode : System.Windows.Forms.TreeNode
        {
            public ViewSheet sheet;
        }

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

        private void RunProgram(object sender, EventArgs e)
        {
            TreeView tree = (TreeView)formWrapper.Controls[0];

            Transaction t = new Transaction(Doc, "Заполнить штампы");
            t.Start();

            foreach (TreeNode building in tree.Nodes)
            {
                foreach (TreeNode tome in building.Nodes)
                {
                    foreach (SheetTreeNode sheetNode in tome.Nodes)
                    {
                        if (!sheetNode.Checked)
                        {
                            continue;
                        }

                        ViewSheet sheet = sheetNode.sheet;

                        for (int i = 1; i<=6; i++)
                        {
                            string paramPos_name = "ADSK_Штамп Строка " + (i).ToString() + " должность";
                            Parameter paramPos = sheet.LookupParameter(paramPos_name);

                            // Проверяем, задана ли должность
                            if (paramPos == null)
                            {
                                continue;
                            }

                            string paramName_name = "ADSK_Штамп Строка " + (i).ToString() + " фамилия";
                            Parameter paramName = sheet.LookupParameter(paramName_name);

                            // Проверяем, задана ли фамилия
                            if (paramName == null)
                            {
                                continue;
                            }

                            FlowLayoutPanel lineWrapper = (FlowLayoutPanel)linesWrapper.Controls[i-1];
                            WinForms.TextBox tb1 = (WinForms.TextBox)lineWrapper.Controls[1];
                            WinForms.TextBox tb2 = (WinForms.TextBox)lineWrapper.Controls[3];

                            // ПРОВЕРИТЬ НА ПЕРЕБИВАНИЕ

                            if ((tb1.Text.Length > 0) || (checkBox.Checked)) paramPos.Set(tb1.Text);
                            if ((tb2.Text.Length > 0) || (checkBox.Checked)) paramName.Set(tb2.Text);
                        }
                    }
                }
            }
            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
