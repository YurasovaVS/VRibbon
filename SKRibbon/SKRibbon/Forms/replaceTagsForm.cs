using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace ReplaceAreaTags
{
    [Transaction(TransactionMode.Manual)]
    public partial class replaceTagsForm : System.Windows.Forms.Form
    {
        Document Doc;
        List<FamilySymbol> roomTagTypes = new List<FamilySymbol>();
        public replaceTagsForm(Document doc)
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

            //Создаем первый заголовок
            Label header1 = new Label();
            header1.Parent = formWrapper;
            formWrapper.Controls.Add(header1);
            header1.Anchor = AnchorStyles.Top;
            header1.Size = new Size(500, 30);
            header1.Text = "Заменить марку типа:";

            //Создаем первый выпадающий список

            System.Windows.Forms.ComboBox comboInitTag = new System.Windows.Forms.ComboBox();
            comboInitTag.Parent = formWrapper;
            formWrapper.Controls.Add(comboInitTag);
            comboInitTag.Anchor = AnchorStyles.Top;
            comboInitTag.Size = new Size(500, 30);
                   
            
            //Создаем второй заголовок
            Label header2 = new Label();
            header2.Parent = formWrapper;
            formWrapper.Controls.Add(header2);
            header2.Anchor = AnchorStyles.Top;
            header2.Size = new Size(500, 30);
            header2.Text = "Заменить на марку типа:";

            //Создаем второй выпадающий список
            System.Windows.Forms.ComboBox comboGoalTag = new System.Windows.Forms.ComboBox();
            comboGoalTag.Parent = formWrapper;
            formWrapper.Controls.Add(comboGoalTag);
            comboGoalTag.Anchor = AnchorStyles.Top;
            comboGoalTag.Size = new Size(500, 30);

            //Заполняем выпадающие списки
            ICollection<Element>roomTagTypesCol = new FilteredElementCollector(doc).
                                                OfCategory(BuiltInCategory.OST_RoomTags).
                                                WhereElementIsElementType().
                                                ToElements();
            foreach (FamilySymbol roomTagType in roomTagTypesCol)
            {
                comboInitTag.Items.Add(roomTagType.Name);
                comboGoalTag.Items.Add(roomTagType.Name);
                roomTagTypes.Add(roomTagType);                
            }
            comboInitTag.SelectedIndex = 0;
            comboGoalTag.SelectedIndex = 1;

            //Создаем третий заголовок
            Label header3 = new Label();
            header3.Parent = formWrapper;
            formWrapper.Controls.Add(header3);
            header3.Anchor = AnchorStyles.Top;
            header3.Size = new Size(500, 30);
            header3.Text = "Выберите листы:";

            //Добавляем чеклист для выбора видов
            CheckedViewList viewList = new CheckedViewList();
            viewList.Parent = formWrapper;
            formWrapper.Controls.Add(viewList);
            viewList.Anchor = AnchorStyles.Top;
            viewList.MinimumSize = new Size(500, 30);
            viewList.Height = 200;
            viewList.CheckOnClick = true;

            //Добавляем виды в чеклист
            IEnumerable<Element> views = new FilteredElementCollector(doc).
                                            OfCategory(BuiltInCategory.OST_Views).
                                            WhereElementIsNotElementType().
                                            ToElements();
            views = views.OrderBy(view => view.Name);
            foreach (Autodesk.Revit.DB.View view in views)
            {
                if (view.ViewType == ViewType.FloorPlan)
                {
                    string viewName = view.Name;
                    if ((viewName != null) & (viewName != ""))
                    {
                        viewList.Items.Add(view.Name);
                        viewList.viewsCollection.Add(view);
                    }
                }
            }

            //Добавляем кнопку
            Button button = new Button();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Width = 200;

            button.Text = "Заменить марки";
            button.Click += ReplaceChosenTags;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;

        }

        //Метод обработки формы
        public void ReplaceChosenTags(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;
            int listIndex = formWrapper.Controls.IndexOf(button) - 1;
            int goalIndex = formWrapper.Controls.IndexOf(button) - 3;
            int initIndex = formWrapper.Controls.IndexOf(button) - 5;

            CheckedViewList viewList = (CheckedViewList)formWrapper.Controls[listIndex];
            System.Windows.Forms.ComboBox initBox = (System.Windows.Forms.ComboBox)formWrapper.Controls[initIndex];
            System.Windows.Forms.ComboBox goalBox = (System.Windows.Forms.ComboBox)formWrapper.Controls[goalIndex];

            foreach (int i in viewList.CheckedIndices)
            {

                Autodesk.Revit.DB.View view = viewList.viewsCollection[i];
                //Выбираем все марки на виде
                ICollection<Element> tags = new FilteredElementCollector(Doc, view.Id).
                                            OfClass(typeof(SpatialElementTag)).
                                            WhereElementIsNotElementType().
                                            ToElements();
                Transaction t = new Transaction(Doc, "Заменить марки");
                t.Start();


                foreach (SpatialElementTag tag in tags)
                {
                    //Если марка нужного нам типа
                    if (tag.GetTypeId() == roomTagTypes[initBox.SelectedIndex].Id)
                    {
                        //Заменяем тип марки
                        tag.ChangeTypeId(roomTagTypes[goalBox.SelectedIndex].Id);
                    }                    
                }
                t.Commit();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }

    //Класс для списка видов
    public class CheckedViewList : CheckedListBox
    {
        public List<Autodesk.Revit.DB.View> viewsCollection;
        public CheckedViewList()
        {
            viewsCollection = new List<Autodesk.Revit.DB.View>();
        }
    }        
}
