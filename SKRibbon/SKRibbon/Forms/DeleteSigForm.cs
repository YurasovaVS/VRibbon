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

            //Добавляем чеклист для выбора видов
            CheckedSheetList sheetList = new CheckedSheetList();
            sheetList.Parent = formWrapper;
            formWrapper.Controls.Add(sheetList);
            sheetList.Anchor = AnchorStyles.Top;
            sheetList.MinimumSize = new Size(500, 30);
            sheetList.Height = 200;
            sheetList.CheckOnClick = true;

            //Добавляем виды в чеклист
            IEnumerable<Element> sheets = new FilteredElementCollector(doc).
                                            OfCategory(BuiltInCategory.OST_Sheets).
                                            WhereElementIsNotElementType().
                                            ToElements();
            sheets = sheets.OrderBy(sheet => sheet.Name);
            foreach (Autodesk.Revit.DB.ViewSheet sheet in sheets)
            {
                string sheetName = sheet.Name;
                if ((sheetName != null) & (sheetName != ""))
                {
                    sheetList.Items.Add(sheet.Name);
                    sheetList.sheetCollection.Add(sheet);
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
            CheckedSheetList sheetList = (CheckedSheetList)formWrapper.Controls[1];
            StringBuilder sb = new StringBuilder();

            Transaction t = new Transaction(Doc, "Убрать подписи");
            t.Start();

            foreach (int j in sheetList.CheckedIndices)
            {
                ViewSheet sheet = sheetList.sheetCollection[j];
                ICollection<Element> signatures = new FilteredElementCollector(Doc, sheet.Id).
                                                OfClass(typeof(ImportInstance)).
                                                ToElements();
                foreach (ImportInstance signature in signatures)
                {
                    string name = signature.LookupParameter("Имя").AsString();
                    sb.AppendLine(name);
                    string pattern = "подпись_*";
                    if (Regex.Match(name, pattern).Success)
                    {
                        sb.AppendLine("Удалена");
                        Doc.Delete(signature.Id);
                    }
                }
                TaskDialog.Show(sheet.Name, sb.ToString());
            }

            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
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
