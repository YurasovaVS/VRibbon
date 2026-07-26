using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static SKRibbon.FormDesign;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace SKRibbon.Forms
{
    public partial class WriteRoomIdToElementParamForm : VForm
    {
        Document Doc;
        ICollection<ElementId> SelectionIds;

        VComboBox TypeOfElementsCB = new VComboBox();
        VComboBox SourceParam = new VComboBox();
        VTextBox TargetParam = new VTextBox();
        VTextBox LinkParam = new VTextBox();
        VTextBox IdParam = new VTextBox();
        VButton OkButton = new VButton();
        int LabelWidth = 200;
        int LabelHeight = 25;

        int ComboWidth = 200;
        int ComboHeight = 25;

        public WriteRoomIdToElementParamForm(Document doc, ICollection<ElementId> selectionIds)
        {
            InitializeComponent();
            Doc = doc;
            SelectionIds = selectionIds;

            // Обертка для всей формы
            FlowLayoutPanel formWrapper = new FlowLayoutPanel();

            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            // Добавляем в форму элементы
            // Строка 1 - выделение или тип
            if (SelectionIds.Count > 0) {
                TypeOfElementsCB.Items.Add("Выделение");
                TypeOfElementsCB.Enabled = false;
            }
            else
            {
                TypeOfElementsCB.Items.Add("Полы");
                TypeOfElementsCB.Items.Add("Потолки");
            }
            TypeOfElementsCB.SelectedIndex = 0;

            System.Windows.Forms.Label selectionLabel = new System.Windows.Forms.Label();
            selectionLabel.Text = "Элементы: ";
            selectionLabel.Size = new System.Drawing.Size(LabelWidth, LabelHeight);
            selectionLabel.Anchor = AnchorStyles.Left;

            TypeOfElementsCB.Size = new System.Drawing.Size(ComboWidth, ComboHeight);
            TypeOfElementsCB.Anchor = AnchorStyles.Left;

            FlowLayoutPanel optionWrapper_1 = new FlowLayoutPanel();
            optionWrapper_1.AutoSize = true;
            optionWrapper_1.FlowDirection = FlowDirection.LeftToRight;

            selectionLabel.Parent = optionWrapper_1;
            optionWrapper_1.Controls.Add(selectionLabel);

            TypeOfElementsCB.Parent = optionWrapper_1;
            optionWrapper_1.Controls.Add(TypeOfElementsCB);

            // Строка 2 - Что пишем
            System.Windows.Forms.Label sourceLabel = new System.Windows.Forms.Label();
            sourceLabel.Text = "Что пишем: ";
            sourceLabel.Size = new System.Drawing.Size(LabelWidth, LabelHeight);
            sourceLabel.Anchor = AnchorStyles.Left;

            SourceParam.Size = new System.Drawing.Size(ComboWidth, ComboHeight);
            SourceParam.Anchor = AnchorStyles.Left;
            SourceParam.Items.Add("Номера помещений");
            SourceParam.Items.Add("Названия помещений");
            SourceParam.SelectedIndex = 0;

            FlowLayoutPanel optionWrapper_2 = new FlowLayoutPanel();
            optionWrapper_2.AutoSize = true;
            optionWrapper_2.FlowDirection = FlowDirection.LeftToRight;

            sourceLabel.Parent = optionWrapper_2;
            optionWrapper_2.Controls.Add(sourceLabel);

            SourceParam.Parent = optionWrapper_2;
            optionWrapper_2.Controls.Add(SourceParam);

            // Строка 3 - Куда пишем
            System.Windows.Forms.Label targetLabel = new System.Windows.Forms.Label();
            targetLabel.Text = "Куда пишем: ";
            targetLabel.Size = new System.Drawing.Size(LabelWidth, LabelHeight);
            targetLabel.Anchor = AnchorStyles.Left;

            TargetParam.Size = new System.Drawing.Size(ComboWidth, ComboHeight);
            TargetParam.Anchor = AnchorStyles.Left;

            FlowLayoutPanel optionWrapper_3 = new FlowLayoutPanel();
            optionWrapper_3.AutoSize = true;
            optionWrapper_3.FlowDirection = FlowDirection.LeftToRight;

            targetLabel.Parent = optionWrapper_3;
            optionWrapper_3.Controls.Add(targetLabel);

            TargetParam.Parent = optionWrapper_3;
            optionWrapper_3.Controls.Add(TargetParam);

            // Строка 4 - Параметр элемента, указывающий на номер помещения
            System.Windows.Forms.Label linkLabel = new System.Windows.Forms.Label();
            linkLabel.Text = "Связь с помещением: ";
            linkLabel.Size = new System.Drawing.Size(LabelWidth, LabelHeight);
            linkLabel.Anchor = AnchorStyles.Left;
            
            ToolTip linkLabelToolTip = new ToolTip();
            linkLabelToolTip.SetToolTip (linkLabel, "Параметр экземляра элемента (пола или потолка), указывающий на номер помещения, которому он принадлежит.");

            LinkParam.Size = new System.Drawing.Size(ComboWidth, ComboHeight);
            LinkParam.Anchor = AnchorStyles.Left;
            LinkParam.Text = "ADSK_Группирование";

            FlowLayoutPanel optionWrapper_4 = new FlowLayoutPanel();
            optionWrapper_4.AutoSize = true;
            optionWrapper_4.FlowDirection = FlowDirection.LeftToRight;

            linkLabel.Parent = optionWrapper_4;
            optionWrapper_4.Controls.Add(linkLabel);

            LinkParam.Parent = optionWrapper_4;
            optionWrapper_4.Controls.Add(LinkParam);

            // Строка 5 - Марка
            System.Windows.Forms.Label idLabel = new System.Windows.Forms.Label();
            idLabel.Text = "Марка: ";
            idLabel.Size = new System.Drawing.Size(LabelWidth, LabelHeight);
            idLabel.Anchor = AnchorStyles.Left;

            IdParam.Size = new System.Drawing.Size(ComboWidth, ComboHeight);
            IdParam.Anchor = AnchorStyles.Left;
            IdParam.Text = "ADSK_Марка";

            FlowLayoutPanel optionWrapper_5 = new FlowLayoutPanel();
            optionWrapper_5.AutoSize = true;
            optionWrapper_5.FlowDirection = FlowDirection.LeftToRight;

            idLabel.Parent = optionWrapper_5;
            optionWrapper_5.Controls.Add(idLabel);

            IdParam.Parent = optionWrapper_5;
            optionWrapper_5.Controls.Add(IdParam);

            // Кнопка
            OkButton.Size = new Size(200, 40);
            OkButton.Anchor = AnchorStyles.Top;
            OkButton.Text = "Начать";
            OkButton.Click += OkButton_Click;

            // Добавляем созданные строки в formwrapper
            optionWrapper_1.Parent = formWrapper;
            formWrapper.Controls.Add(optionWrapper_1);

            optionWrapper_2.Parent = formWrapper;
            formWrapper.Controls.Add(optionWrapper_2);

            optionWrapper_3.Parent = formWrapper;
            formWrapper.Controls.Add(optionWrapper_3);

            optionWrapper_4.Parent = formWrapper;
            formWrapper.Controls.Add(optionWrapper_4);

            optionWrapper_5.Parent = formWrapper;
            formWrapper.Controls.Add(optionWrapper_5);

            OkButton.Parent = formWrapper;
            formWrapper.Controls.Add(OkButton);

            this.AutoSize = true;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            ICollection<ElementId> elementIds;
            switch (TypeOfElementsCB.Text)
            {
                case "Выделение":
                    elementIds = SelectionIds;
                    break;
                case "Потолки":
                    elementIds = new FilteredElementCollector(Doc).
                        OfCategory(BuiltInCategory.OST_Ceilings).
                        WhereElementIsNotElementType().
                        ToElementIds();
                    break;
                default: // по дефолту будут полы
                    elementIds = new FilteredElementCollector(Doc).
                        OfCategory(BuiltInCategory.OST_Floors).
                        WhereElementIsNotElementType().
                        ToElementIds();
                    break;
            }

            ICollection<Element> rooms = new FilteredElementCollector(Doc).
                        OfCategory(BuiltInCategory.OST_Rooms).
                        WhereElementIsNotElementType().
                        ToElements();

            Dictionary<string, HashSet<string>>roomNamesByType = new Dictionary<string, HashSet<string>>();
            Dictionary<string, Element>typesByMark = new Dictionary<string, Element>();

            // Начало обработки и сортировки элементов по типам
            foreach (ElementId elementId in elementIds) {
                Element el = Doc.GetElement(elementId);
                Element elType = Doc.GetElement(el.GetTypeId());

                Parameter uniqueTypeIdParam = elType.LookupParameter(IdParam.Text);
                Parameter elementLinkParam = el.LookupParameter(LinkParam.Text);

                if (uniqueTypeIdParam != null) {
                    string uniqueTypeIdText = uniqueTypeIdParam.AsString();
                    if ((uniqueTypeIdText != "") && (uniqueTypeIdText != null))
                    {
                        // Если такого типа еще не было, добавляем его в словари
                        if (!roomNamesByType.ContainsKey(uniqueTypeIdText))
                        {
                            roomNamesByType.Add(uniqueTypeIdText, new HashSet<string>());
                            typesByMark.Add(uniqueTypeIdText, elType);
                        }
                        
                        if (elementLinkParam != null)
                        {
                            string elementLinkText = elementLinkParam.AsString();
                            if ((elementLinkText != "") && (elementLinkText != null))
                            {
                                // Ищем комнату, к которой привязан данный элемент
                                foreach (Element room in rooms)
                                {
                                    Room trueRoom = room as Room;
                                    if (trueRoom.Number.ToString() == elementLinkText)
                                    {
                                        switch (SourceParam.Text)
                                        {
                                            case ("Номера помещений"): roomNamesByType[uniqueTypeIdText].Add(trueRoom.Number.ToString()); 
                                                break;
                                            default:
                                                Parameter roomNameParam = room.get_Parameter(BuiltInParameter.ROOM_NAME);
                                                if (roomNameParam != null)
                                                {
                                                    string roomName = roomNameParam.AsString();
                                                    if ((roomName != "") && (roomName != null))
                                                    {
                                                        roomNamesByType[uniqueTypeIdText].Add(roomName);
                                                    }
                                                }
                                                    
                                                break;
                                        }
                                        break;
                                    }
                                } // Конец перебора комнат
                            }
                        }

                    }
                }
            } // Конец обработки элементов

            Transaction t = new Transaction(Doc, "Прописать наименования или номера в тип");
            t.Start();
            // Начало обработки словарей
            foreach (string key in roomNamesByType.Keys)
            {
                List <string> tempOrderedList = roomNamesByType[key].OrderBy(x => x).ToList();
                string line = string.Join(", ", tempOrderedList);
                Parameter targetParam = typesByMark[key].LookupParameter(TargetParam.Text);
                if (targetParam != null)
                {
                    targetParam.Set(line);
                }
            }
            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
