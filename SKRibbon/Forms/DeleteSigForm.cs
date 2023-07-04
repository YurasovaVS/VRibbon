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
            buildingsDict = FormUtils.CollectSheetDictionary(doc, true);
            TreeView sheetTree = FormUtils.CreateSheetTreeView(buildingsDict);

            sheetTree.Parent = formWrapper;
            formWrapper.Controls.Add(sheetTree);
            sheetTree.Anchor = AnchorStyles.Top;
            sheetTree.MinimumSize = new Size(500, 30);
            sheetTree.Height = 200;
            sheetTree.CheckBoxes = true;
            sheetTree.AfterCheck += node_AfterCheck;

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
                    foreach (FormUtils.SheetTreeNode sheetNode in tome.Nodes)
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
