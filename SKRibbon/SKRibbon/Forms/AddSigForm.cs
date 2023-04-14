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
            CheckedSheetList sheetList = (CheckedSheetList)formWrapper.Controls[3];
            StringBuilder sb = new StringBuilder();

            Transaction t = new Transaction(Doc, "Вставить подписи");
            t.Start();

            foreach (int j in sheetList.CheckedIndices)
            {
                ViewSheet sheet = sheetList.sheetCollection[j];
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
                }
            }
            if (sb.Length == 0)
            {
                sb.AppendLine("Подписи проставлены. Ошибок нет.");
            }
            TaskDialog.Show("Ошибки", sb.ToString());

            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
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
}
