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
        public NewTotalForm(Document doc)
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
            header.Text = "Введите новую общую площадь по каждому объекту";
            
            //Находим все помещения в проекте
            ICollection<Element> rooms = new FilteredElementCollector(doc).
                                    OfClass(typeof(SpatialElement)).
                                    WhereElementIsNotElementType().
                                    ToElements();

            foreach (SpatialElement room in rooms)
            {
                Parameter paramArea = room.LookupParameter("Площадь");
                Parameter paramBuilding = room.LookupParameter("ADSK_Номер здания");
                Parameter paramRole = room.LookupParameter("Назначение");

                if (paramRole != null)
                {
                    if (paramRole.AsString() == "Квартиры")
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
                lineWrapper.Parent = formWrapper;
                formWrapper.Controls.Add(lineWrapper);
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

            //Добавляем кнопку
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

        public class CheckedViewList : CheckedListBox {
            public List<Autodesk.Revit.DB.View> viewsCollection;
            public CheckedViewList()
            {
                viewsCollection = new List<Autodesk.Revit.DB.View> ();
            }
        }
        
        
        public void CalculateAdjustedAreas (object sender, EventArgs e)
        {
            Button button = (Button)sender;
            FlowLayoutPanel formWrapper = (FlowLayoutPanel)button.Parent;
                                          

            //Собираем и парсим новые площади
            foreach (FlowLayoutPanel panel in formWrapper.Controls.OfType<FlowLayoutPanel>())
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

            //Выбираем все комнаты В ПРОЕКТЕ
            ICollection<Element> rooms = new FilteredElementCollector(Doc).
                                        OfClass(typeof(SpatialElement)).
                                        WhereElementIsNotElementType().
                                        ToElements();

            Transaction t = new Transaction(Doc, "Подогнать площади");
            t.Start();

            foreach (SpatialElement room in rooms)
            {
                Parameter paramArea = room.LookupParameter("Площадь");
                Parameter paramBuilding = room.LookupParameter("ADSK_Номер здания");
                Parameter paramRole = room.LookupParameter("Назначение");

                if (paramRole != null)
                {
                    if (paramRole.AsString() == "Квартиры")
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

                            Parameter paramNewArea = room.LookupParameter("Комментарии");
                            if (paramNewArea != null)
                            {
                                paramNewArea.Set(newArea.ToString());
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

                            Parameter paramNewArea = room.LookupParameter("Комментарии");
                            if (paramNewArea != null)
                            {
                                paramNewArea.Set(newArea.ToString());
                            }
                        }
                    }
                }

            }

            t.Commit();
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void NewTotal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar)) return;
            if (Char.IsControl(e.KeyChar)) return;
            if ((e.KeyChar == ',') && ((sender as System.Windows.Forms.TextBox).Text.Contains(',') == false)) return;
            if ((e.KeyChar == ',') && ((sender as System.Windows.Forms.TextBox).SelectionLength == (sender as System.Windows.Forms.TextBox).TextLength)) return;
            e.Handled = true;
        }
    }
}
