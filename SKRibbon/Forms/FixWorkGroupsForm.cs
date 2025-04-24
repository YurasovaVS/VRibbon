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
using System.Collections;
using System.Net.NetworkInformation;

namespace SKRibbon
{
    
    public partial class FixWorkGroupsForm : WinForms.Form
    {
        Dictionary<BuiltInCategory, string> CategoryNames = new Dictionary<BuiltInCategory, string>()
        {
            { BuiltInCategory.OST_CurtaSystem, "Витражи"},
            { BuiltInCategory.OST_Doors, "Двери"},
            { BuiltInCategory.OST_Roofs, "Кровля"},
            { BuiltInCategory.OST_Stairs, "Лестницы"},
            { BuiltInCategory.OST_PlumbingFixtures, "Оборудование сантех."},
            { BuiltInCategory.OST_DuctTerminal, "Оборудование сантех."},
            { BuiltInCategory.OST_MechanicalEquipment, "Оборудование мех."},
            { BuiltInCategory.OST_Railings, "Ограждения"},
            { BuiltInCategory.OST_StairsRailing, "Ограждения"},
            { BuiltInCategory.OST_Windows, "Окна"},
            { BuiltInCategory.OST_ShaftOpening, "Отверстия"},
            { BuiltInCategory.OST_Grids, "Оси"},
            { BuiltInCategory.OST_Floors, "Полы"},
            { BuiltInCategory.OST_Rooms, "Помещения"},
            { BuiltInCategory.OST_RoomSeparationLines, "Помещения"},
            { BuiltInCategory.OST_Ceilings, "Потолки"},
            { BuiltInCategory.OST_Levels, "Уровни"},
            { BuiltInCategory.OST_Mass, "Формы"}
        };

        //Dictionary<string, BuiltInCategory> NamesCategory = new Dictionary<string, BuiltInCategory>();
        HashSet<string> NamesCategory = new HashSet<string>();

        Dictionary<string, Workset> Name_Workset = new Dictionary<string, Workset>();
        HashSet<string> Errors = new HashSet<string>();

        Document Doc;

        WinForms.FlowLayoutPanel formWrapper = new WinForms.FlowLayoutPanel();
        WinForms.FlowLayoutPanel CategoriesWrapper = new WinForms.FlowLayoutPanel();
        WinForms.FlowLayoutPanel CategoriesByTypeWrapper = new WinForms.FlowLayoutPanel();

        int LabelWidth = 150;
        int LabelHeight = 20;

        int TextBoxWidth = 150;
        int TextBoxHeight = 20;

        int ComboBoxWidth = 200;
        int ComboBoxHeight = 20;

        public FixWorkGroupsForm(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            this.AutoScroll = true;
            this.Width = 710;
            this.Height = 480;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle;

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            CategoriesByTypeWrapper.AutoSize = true;
            CategoriesByTypeWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            CategoriesByTypeWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(CategoriesByTypeWrapper);

            // Добавляем строку для КЖ (Label, TextBox, ComboBox);

            FlowLayoutPanel ConstrFLP = WorkgroupAndFamilyFLP("КЖ", "КЖ");
            FlowLayoutPanel DisabledPeopleFLP = WorkgroupAndFamilyFLP("ОДИ", "ОДИ");
            FlowLayoutPanel WallsInternalFLP = WorkgroupAndFamilyFLP("Внутренние стены", "АР_В");
            FlowLayoutPanel WallsExternalFLP = WorkgroupAndFamilyFLP("Наружные стены", "АР_Н");
            FlowLayoutPanel WallsDesignFLP = WorkgroupAndFamilyFLP("Отделка", "АР_О");
            FlowLayoutPanel CurtainWalls = WorkgroupAndFamilyFLP("Витражи", "витраж");
            FlowLayoutPanel GenPlan = WorkgroupAndFamilyFLP("Генплан", "ГП");



            IEnumerable<Workset> worksetsIE = new FilteredWorksetCollector(Doc).OfKind(WorksetKind.UserWorkset).ToWorksets();

            int i = 0;
            foreach (Workset workset in worksetsIE) {
                Name_Workset.Add(workset.Name, workset);

                // Заполняем существующие выпадающие списки
                foreach (var panel in CategoriesByTypeWrapper.Controls)
                {
                    FlowLayoutPanel p = panel as FlowLayoutPanel;
                    WinForms.ComboBox box = p.Controls[2] as WinForms.ComboBox;
                    WinForms.Label label = p.Controls[0] as WinForms.Label;
                    box.Items.Add(workset.Name);
                    var words = label.Text.Split(' ');
                    if (workset.Name.Contains(words[0])) box.SelectedIndex = i;
                }
                i++;
            }

            // Добавляем блок с категориями

            CategoriesWrapper.FlowDirection = WinForms.FlowDirection.TopDown;
            CategoriesWrapper.AutoSize = true;

            foreach (KeyValuePair<BuiltInCategory, string> category in CategoryNames) {
                if (NamesCategory.Add(category.Value))
                {
                    WinForms.FlowLayoutPanel catPanel = WorkgroupFLP(category.Value);
                    catPanel.Parent = CategoriesWrapper;
                    CategoriesWrapper.Controls.Add(catPanel);
                }                
            }

            formWrapper.Controls.Add(CategoriesWrapper);
            CategoriesWrapper.Parent = formWrapper;

            // Добавить кнопку

            WinForms.Button OkButton = new WinForms.Button();
            OkButton.Size = new Size(200, 50);
            OkButton.Text = "ОК";
            OkButton.Click += RunFixing;

            formWrapper.Controls.Add(OkButton);
            OkButton.Parent = formWrapper;

            //ICollection<Element> elements = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements();

        }

        private void RunFixing(object sender, EventArgs e)
        {
            ICollection<Element> elements = new FilteredElementCollector(Doc).WhereElementIsNotElementType().ToElements();
            StringBuilder sb_ReadOnly = new StringBuilder();
            StringBuilder sb_TypeNotSupported = new StringBuilder();
            StringBuilder sb_ElementHasNoCategory = new StringBuilder();
            sb_ReadOnly.AppendLine("Рабочий набор данных элементов стоит в режиме \"Только для чтения\": ");
            sb_TypeNotSupported.AppendLine("Тип данных элементов не поддерживается. Обратитесь к разработчику:");
            sb_ElementHasNoCategory.AppendLine("У данных элементов отсутствует категория: ");

            Transaction t = new Transaction(Doc, "Исправить рабочие наборы");
            t.Start();

            foreach (Element element in elements) {

                if (element.Id.ToString() == "1212844")
                {
                    bool flag = true;
                }

                // Проверяем, существует ли такой параметр, и не ReadOnly ли он
                Parameter worksetParam = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                if (worksetParam == null) continue;
                if (worksetParam.IsReadOnly) {
                    sb_ReadOnly.AppendLine(element.Name + " " + element.Id);
                    continue;
                }
                

                string worksetName = "";

                bool skipFurtherChecks = false;

                // Проверяем, не относится ли он к рабочим наборам по типу (КЖ, ОДИ, внутренние и внешние стены, геплан)
                foreach (var catByTypeWrapper in CategoriesByTypeWrapper.Controls)
                {                    
                    if (skipFurtherChecks) continue;

                    FlowLayoutPanel panel = catByTypeWrapper as FlowLayoutPanel;
                    WinForms.TextBox text = panel.Controls[1] as WinForms.TextBox;
                    WinForms.ComboBox combo = panel.Controls[2] as WinForms.ComboBox;                                     

                    if (element.Name.Contains(text.Text)) {
                        worksetName = combo.Text;
                        skipFurtherChecks = true;
                    }
                }
                if (!skipFurtherChecks)
                {
                    // Проверяем, есть ли у данного элемента категория
                    if (element.Category == null)
                    {
                        sb_ElementHasNoCategory.AppendLine(element.Name + " " + element.Id);
                        continue;
                    }
                    // Проверяем, есть ли такая категория в нашем словаре
                    if (!CategoryNames.ContainsKey((BuiltInCategory)element.Category.Id.Value))
                    {
                        sb_TypeNotSupported.AppendLine(element.Name + " " + element.Id + (BuiltInCategory)element.Category.Id.IntegerValue);
                        continue;
                    }
                    // Смотрим, к какому рабочему набору относится элемент (BuiltInCategory)
                    foreach (var categoryFLP in CategoriesWrapper.Controls)
                    {
                        if (skipFurtherChecks) continue; // Если мы уже перенесли элемент в какой-либо рабочий набор,              


                        FlowLayoutPanel panel = categoryFLP as FlowLayoutPanel;
                        WinForms.Label catName = panel.Controls[0] as WinForms.Label;
                        WinForms.ComboBox combo = panel.Controls[1] as WinForms.ComboBox;

                        if (CategoryNames[(BuiltInCategory)element.Category.Id.Value] != catName.Text) continue;
                        worksetName = combo.Text;
                        skipFurtherChecks = true;
                    }
                }                
                if (worksetName == "") continue; // Если по какой-то причине worksetName остается не задан, пропускаем элемент
                worksetParam.Set(Name_Workset[worksetName].Id.IntegerValue);
            } // Конец цикла foreach (Element element in elements)

            t.Commit();

            TaskDialog.Show("Ошибки", sb_ReadOnly.ToString());
            TaskDialog.Show("Ошибки", sb_TypeNotSupported.ToString());
            // TaskDialog.Show("Ошибки", sb_ElementHasNoCategory.ToString());           
                

            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
        }

        // Создание группы для проверки по категории
        public WinForms.FlowLayoutPanel WorkgroupFLP(string Name)
        {
            WinForms.FlowLayoutPanel workgroupFLP = new WinForms.FlowLayoutPanel();
            WinForms.Label workgroupName = new WinForms.Label();
            WinForms.ComboBox workgroupsCB = new WinForms.ComboBox();

            workgroupFLP.AutoSize = true;
            workgroupFLP.FlowDirection = WinForms.FlowDirection.LeftToRight;

            workgroupFLP.Controls.Add(workgroupName);
            workgroupFLP.Controls.Add(workgroupsCB);
            workgroupName.Parent = workgroupFLP;
            workgroupsCB.Parent = workgroupFLP;

            workgroupName.Text = Name;
            workgroupName.Size = new Size(LabelWidth, LabelHeight);

            workgroupsCB.Size = new Size(ComboBoxWidth, ComboBoxHeight);
            foreach (KeyValuePair<string, Workset> workset in Name_Workset) {
                int i = workgroupsCB.Items.Add(workset.Key);
                var words = Name.Split(' ');
                if (workset.Key.Contains(words[0])) workgroupsCB.SelectedIndex = i;
            }            
            return workgroupFLP;
        }

        // Создание группы для проверки по семейству
        public WinForms.FlowLayoutPanel WorkgroupAndFamilyFLP(string Name, string TextBoxText)
        {
            WinForms.TextBox FLPTextBox = new WinForms.TextBox();
            WinForms.ComboBox FLPComboBox = new WinForms.ComboBox();

            WinForms.FlowLayoutPanel FLPanel = new WinForms.FlowLayoutPanel();
            FLPanel.FlowDirection = WinForms.FlowDirection.LeftToRight;
            FLPanel.AutoSize = true;


            WinForms.Label ConstrLabel = new WinForms.Label();
            ConstrLabel.Size = new Size(LabelWidth, LabelHeight);
            ConstrLabel.Text = Name;

            FLPTextBox.Size = new Size(TextBoxWidth, TextBoxHeight);
            FLPTextBox.Text = TextBoxText;

            FLPComboBox.Size = new Size(ComboBoxWidth, ComboBoxHeight);

            FLPanel.Controls.Add(ConstrLabel);
            FLPanel.Controls.Add(FLPTextBox);
            FLPanel.Controls.Add(FLPComboBox);

            ConstrLabel.Parent = FLPanel;
            FLPTextBox.Parent = FLPanel;
            FLPComboBox.Parent = FLPanel;

            CategoriesByTypeWrapper.Controls.Add(FLPanel);
            FLPanel.Parent = CategoriesByTypeWrapper;

            return FLPanel;
        }
    }   
}
