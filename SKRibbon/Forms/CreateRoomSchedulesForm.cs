using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinForm = System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;

namespace SKRibbon
{
    [Transaction(TransactionMode.Manual)]
    public partial class CreateRoomSchedulesForm : WinForm.Form 
    {
        Document Doc;
        public CreateRoomSchedulesForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            this.AutoSize = true;
            this.AutoScroll = true;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            //Создаем Wrapper для содержимого формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;
            formWrapper.BorderStyle = BorderStyle.FixedSingle;
            formWrapper.Padding = new Padding(5, 5, 5, 5);

            //Создаем первый заголовок
            Label header1 = new Label();
            header1.Parent = formWrapper;
            formWrapper.Controls.Add(header1);
            header1.Anchor = AnchorStyles.Top;
            header1.Size = new Size(500, 30);
            header1.Text = "Выберите уровни:";

            //Создаем список уровней

            CheckedListBox levelsListBox = new CheckedListBox();
            levelsListBox.Parent = formWrapper;
            formWrapper.Controls.Add(levelsListBox);
            levelsListBox.Anchor = AnchorStyles.Top;
            levelsListBox.Size = new Size(400, 400);

            ICollection<Element> levels = new FilteredElementCollector(doc).
                                                OfCategory(BuiltInCategory.OST_Levels).
                                                WhereElementIsNotElementType().
                                                ToElements();
            foreach (Element level in levels) { 
                levelsListBox.Items.Add(level.Name);
            }

            // Добавить кнопку "Инвертировать выделение"
            Button invertButton = new Button();
            invertButton.Parent = formWrapper;
            formWrapper.Controls.Add(invertButton);
            invertButton.Anchor = AnchorStyles.Top;
            invertButton.Size = new Size(150, 30);
            invertButton.Text = "Инвертировать выделение";

            // Добавить кнопку "Создать спецификации"


        }
    }
}
