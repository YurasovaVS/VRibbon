using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Xml.Linq;

namespace FakeArea
{
    [Transaction(TransactionMode.Manual)]
    public partial class NewTotalForm : System.Windows.Forms.Form
    {
        Document Doc;
        Dictionary<string, BuildingAdjustments> dictionary = new Dictionary<string, BuildingAdjustments>();
        ICollection<Element> rooms;
        string currentRoomPurpose;
        public NewTotalForm(Document doc)
        {
            InitializeComponent();
            
            Doc = doc;
            rooms = new FilteredElementCollector(doc).
                                    OfClass(typeof(SpatialElement)).
                                    WhereElementIsNotElementType().
                                    ToElements();
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

            //Создаем заголовок выпадающего списка. Индекс во wrapper'е: [0]
            Label comboHeader = new Label();
            comboHeader.Parent = formWrapper;
            formWrapper.Controls.Add(comboHeader);
            comboHeader.Anchor = AnchorStyles.Top;
            comboHeader.Size = new Size(500, 30);
            comboHeader.Text = "Выберите назначение помещений";

            //Создаем выпадающий список. Индекс во wrapper'е: [1]
            System.Windows.Forms.ComboBox comboBox = new System.Windows.Forms.ComboBox();
            comboBox.Parent = formWrapper;
            formWrapper.Controls.Add(comboBox);
            comboBox.Anchor = AnchorStyles.Top;
            comboBox.Size = new Size(500, 30);

            //Заполняем выпадающий список
            HashSet<string> roomPurposes = new HashSet<string>();

            foreach (SpatialElement room in rooms)
            {
                Parameter param = room.LookupParameter("Назначение");
                if (param != null)
                {
                    
                    if ((param.AsString() == "") || (param.AsString() == null))
                    {
                        currentRoomPurpose = "<НЕ ОПРЕДЕЛЕНО>";
                    }
                    else
                    {
                        currentRoomPurpose = param.AsString();
                    }
                    bool flag = roomPurposes.Add(currentRoomPurpose);
                    if (flag)
                    {
                        comboBox.Items.Add(currentRoomPurpose);
                    }
                }
            }
            comboBox.SelectedIndex = 0;
            currentRoomPurpose = comboBox.Items[0].ToString();
            comboBox.SelectedValueChanged += OnSelectionChanged;

            //Создаем заголовок. Индекс во wrapper'е: [2]
            Label header = new Label();
            header.Parent = formWrapper;
            formWrapper.Controls.Add(header);
            header.Anchor = AnchorStyles.Top;
            header.Size = new Size(500, 30);
            header.Text = "Введите новую общую площадь по каждому объекту";

            //Создаем wrapper для элементов зданий. Индекс во wrapper'е: [3]
            FlowLayoutPanel buildingsWrapper = new FlowLayoutPanel();
            buildingsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(buildingsWrapper);

            buildingsWrapper.FlowDirection = FlowDirection.TopDown;
            buildingsWrapper.AutoSize = true;
            buildingsWrapper.BorderStyle = BorderStyle.FixedSingle;
            buildingsWrapper.Padding = new Padding(5, 5, 5, 5);

            RecalculateAreas(currentRoomPurpose, buildingsWrapper);
       
            //Добавляем кнопку. Индекс во wrapper'е: [4]
            Button button = new Button();
            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
            button.Width = 200;

            button.Text = "Пересчитать площади";
            button.Click += CalculateAdjustedAreas;

            this.Width = formWrapper.Width + 5;
            this.Height = formWrapper.Height + 5;
        }

        //----------------------------------------------------
        // Классы
        public class BuildingAdjustments {
            public double totalArea;
            public double newTotal;
            public double adjuster;
            public BuildingAdjustments()
            {
                totalArea = 0;
                newTotal = 0;
                adjuster = 0;
            }
        }
       
        //----------------------------------------------------
        // События
        public void OnSelectionChanged (object sender, EventArgs e)
        {
            System.Windows.Forms.ComboBox comboBox = (System.Windows.Forms.ComboBox)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)comboBox.Parent;
            FlowLayoutPanel buildingsWrapper = (FlowLayoutPanel)formWrapper.Controls[3];
            currentRoomPurpose = comboBox.SelectedItem.ToString();
            RecalculateAreas(currentRoomPurpose, buildingsWrapper);
        }

        public void CalculateAdjustedAreas (object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;
            FlowLayoutPanel buildingsWrapper = (FlowLayoutPanel)formWrapper.Controls[3];                                          

            //Собираем и парсим новые площади
            foreach (FlowLayoutPanel panel in buildingsWrapper.Controls.OfType<FlowLayoutPanel>())
            {
                Label buildingName = (Label)panel.Controls[0];
                System.Windows.Forms.TextBox newTotal = (System.Windows.Forms.TextBox)panel.Controls[1];
                dictionary[buildingName.Text].newTotal = double.Parse(newTotal.Text);
            }
            // Считаем коэффициент правки для каждого здания
            foreach (var building in dictionary)
            {
                building.Value.adjuster = (building.Value.newTotal - building.Value.totalArea) / building.Value.totalArea;
            }  

            Transaction t = new Transaction(Doc, "Подогнать площади");
            t.Start();

            // Заполняем Комментарий помещения
            foreach (SpatialElement room in rooms)
            {
                Parameter paramArea = room.LookupParameter("Площадь");
                Parameter paramBuilding = room.LookupParameter("ADSK_Номер здания");
                Parameter paramRole = room.LookupParameter("Назначение");
                Parameter paramComment = room.LookupParameter("Комментарии");

                if (paramRole != null)
                {
                    if (paramRole.AsString() == currentRoomPurpose)
                    {

                        // Определяем, к какому зданию относится комната
                        string bName = "NONE-NONE";
                        if (paramBuilding != null)
                        {
                            bName = paramBuilding.AsString();
                        }
                        if ((bName == "") | (bName == null))
                        {
                            bName = "NONE-NONE";
                        }

                        // Определяем новую площадь и вставляем в проект
                        if (paramArea != null)
                        {
                            double newArea = paramArea.AsDouble() * (1 + dictionary[bName].adjuster);
                            newArea = UnitUtils.ConvertFromInternalUnits(newArea, UnitTypeId.SquareMeters);
                            newArea = Math.Round(newArea, 2);
                                                        
                            if (paramComment != null)
                            {
                                paramComment.Set(newArea.ToString());
                            }
                        }
                    }
                    else 
                    {
                        if (paramArea != null)
                        {
                            double newArea = paramArea.AsDouble();
                            newArea = UnitUtils.ConvertFromInternalUnits(newArea, UnitTypeId.SquareMeters);
                            newArea = Math.Round(newArea, 2);

                            if (paramComment != null)
                            {
                                string temp = paramComment.AsString();
                                if ((temp == "") | (temp == null))
                                {
                                    paramComment.Set(newArea.ToString());
                                }
                            }
                            else
                            {
                                TaskDialog.Show("Debug", "Что-то пошло не так");
                            }                            
                        }
                    }
                }
            } // Конец foreach (room in rooms)

            t.Commit();
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        } // Конец функции CalculateAdjustedAreas

        // Функция, запрещающая вводить в поля все, кроме цифр и точки
        private void NewTotal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar)) return;
            if (Char.IsControl(e.KeyChar)) return;
            if ((e.KeyChar == ',') && ((sender as System.Windows.Forms.TextBox).Text.Contains(',') == false)) return;
            if ((e.KeyChar == ',') && ((sender as System.Windows.Forms.TextBox).SelectionLength == (sender as System.Windows.Forms.TextBox).TextLength)) return;
            e.Handled = true;
        }

        //-------------------------------------------------------------
        // Функции

        // Функция заполнения формы, в зависимости от назначения здания
        private void RecalculateAreas (string purpose, FlowLayoutPanel wrapper)
        {
            dictionary.Clear();
            wrapper.Controls.Clear();
            
            foreach (SpatialElement room in rooms)
            {
                Parameter paramArea = room.LookupParameter("Площадь");
                Parameter paramBuilding = room.LookupParameter("ADSK_Номер здания");
                Parameter paramRole = room.LookupParameter("Назначение");

                if (paramRole != null)
                {
                    if (paramRole.AsString() == purpose)
                    {
                        //Если помещение не относится ни к одному зданию, задаем ему здание NONE-NONE
                        string buildingName = "NONE-NONE";
                        if (paramBuilding != null)
                        {
                            buildingName = paramBuilding.AsString();
                        }
                        if ((buildingName == "") | (buildingName == null))
                        {
                            buildingName = "NONE-NONE";
                        }

                        //Если такое здание еще не встречалось в словаре, добавляем его в словарь
                        if (!dictionary.ContainsKey(buildingName))
                        {
                            dictionary.Add(buildingName, new BuildingAdjustments());
                        }

                        //Прибавляем площадь комнаты в общую площадь соответствующего здания
                        if (paramArea != null)
                        {
                            double roomArea = paramArea.AsDouble();
                            dictionary[buildingName].totalArea += roomArea;
                        }
                    }
                }
            }

            foreach (var building in dictionary)
            {
                // Переводим текующую площадь из квадратных футов в квадратные метры
                building.Value.totalArea = UnitUtils.ConvertFromInternalUnits(building.Value.totalArea, UnitTypeId.SquareMeters);

                // Создаем wrapper для строчки со вводом новой площади здания
                FlowLayoutPanel lineWrapper = new FlowLayoutPanel();
                lineWrapper.FlowDirection = FlowDirection.LeftToRight;
                lineWrapper.AutoSize = true;
                lineWrapper.Parent = wrapper;
                wrapper.Controls.Add(lineWrapper);
                lineWrapper.Anchor = AnchorStyles.Top;
                lineWrapper.BorderStyle = BorderStyle.FixedSingle;

                // Заголовок - имя здания
                Label buildingName = new Label();
                buildingName.Parent = lineWrapper;
                lineWrapper.Controls.Add(buildingName);
                buildingName.Size = new Size(100, 30);
                buildingName.Text = building.Key.ToString();
                buildingName.Anchor = AnchorStyles.Left;

                // Поле для ввода новой площади
                System.Windows.Forms.TextBox newTotal = new System.Windows.Forms.TextBox();
                newTotal.Parent = lineWrapper;
                lineWrapper.Controls.Add(newTotal);
                newTotal.Size = new Size(300, 30);
                double tA = Math.Round(building.Value.totalArea, 2);
                newTotal.Text = tA.ToString();
                newTotal.KeyPress += NewTotal_KeyPress;
            }
        } // Конец функции RecalculateAreas
    } // Конец класса NewTotalForm
} // Конец пространства имен FakeArea
