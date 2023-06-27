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
using System.Text.RegularExpressions;

namespace SheetRenamer
{
    [Transaction(TransactionMode.Manual)]
    public partial class RenameSheetsForm : System.Windows.Forms.Form
    {
        Document Doc;
        Dictionary<string, Dictionary<string, List<ViewSheet>>> buildingsDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
        System.Windows.Forms.ComboBox buildingsCB = new System.Windows.Forms.ComboBox();
        System.Windows.Forms.ComboBox tomesCB = new System.Windows.Forms.ComboBox();
        System.Windows.Forms.ComboBox sheetsCB = new System.Windows.Forms.ComboBox();

        System.Windows.Forms.TextBox prefixField = new System.Windows.Forms.TextBox();
        System.Windows.Forms.TextBox numField = new System.Windows.Forms.TextBox();
        System.Windows.Forms.TextBox suffixField = new System.Windows.Forms.TextBox();
        System.Windows.Forms.TextBox lengthOfNumField = new System.Windows.Forms.TextBox();

        HashSet<string> sheetNums = new HashSet<string>();        

        public RenameSheetsForm(Document doc)
        {
            Doc = doc;
            this.AutoScroll = true;
            this.Width = 600;
            InitializeComponent();

            FlowLayoutPanel formWrapper = new FlowLayoutPanel();
            formWrapper.AutoSize = true;
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

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
            Dictionary<string, Dictionary<string, List<ViewSheet>>> tempDict = new Dictionary<string, Dictionary<string, List<ViewSheet>>>();
            foreach (ViewSheet sheet in sheets)
            {
                sheetNums.Add(sheet.SheetNumber);
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
                    if (!tempDict.ContainsKey(building))
                    {
                        Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                        tempDict.Add(building, tomesDict);
                    }
                    // Если во вложенном словаре здания нет такого тома, создаем его
                    if (!tempDict[building].ContainsKey(tome))
                    {
                        List<ViewSheet> sheetsList = new List<ViewSheet>();
                        tempDict[building].Add(tome, sheetsList);
                    }
                    // Добавляем лист в нужный том
                    tempDict[building][tome].Add(sheet);
                }
            } // Конец создания словаря листов

            // ОТСОРТИРОВАТЬ ВСЕ СПИСКИ
            foreach (var building in tempDict)
            {
                Dictionary<string, List<ViewSheet>> tomesDict = new Dictionary<string, List<ViewSheet>>();
                buildingsDict.Add(building.Key, tomesDict);
                foreach (var tome in building.Value)
                {
                    List<ViewSheet> sheetsList = tome.Value.OrderBy(sheet => sheet.SheetNumber).ToList();
                    buildingsDict[building.Key].Add(tome.Key, sheetsList);     
                }
            }

            // Создание заголовков
            Label buildingsHeader = new Label();
            Label tomesHeader = new Label();
            Label sheetHeader = new Label();

            buildingsHeader.Size = new System.Drawing.Size(500, 30);
            tomesHeader.Size = new System.Drawing.Size(500, 30);
            sheetHeader.Size = new System.Drawing.Size(500, 30);

            buildingsHeader.Text = "Выберите здание";
            tomesHeader.Text = "Выберите раздел";
            sheetHeader.Text = "Выберите первый лист для перенумерации";



            // Инициализация выпадающего списка со зданиями
            // Инициализация выпадающего списка с томами
            // Инициализация выпадающего списка с листами

            // Размеры
            buildingsCB.Size = new System.Drawing.Size(400, 30);
            tomesCB.Size = new System.Drawing.Size(400, 30);
            sheetsCB.Size = new System.Drawing.Size(400, 30);

            // Заполняем
            foreach (var building in buildingsDict)
            {
                buildingsCB.Items.Add(building.Key);
            }
            buildingsCB.SelectedIndex = 0;

            string buildingsCurrentKey = buildingsCB.SelectedItem.ToString();
            RepopulateComboBox(tomesCB, buildingsDict[buildingsCurrentKey]);
            string tomesCurrentKey = tomesCB.SelectedItem.ToString();
            RepopulateComboBox(sheetsCB, buildingsDict[buildingsCurrentKey][tomesCurrentKey]);

            // Прикручиваем методы для смены содержимого выпадающих списков

            buildingsCB.DropDownClosed += OnBuildingChanged;
            tomesCB.DropDownClosed += OnTomeChanged;
            sheetsCB.DropDownClosed += OnSheetChanged;

            // Инициализируем заголовки полей
            FlowLayoutPanel headersWrapper = new FlowLayoutPanel();
            headersWrapper.FlowDirection = FlowDirection.LeftToRight;
            headersWrapper.AutoSize = true;

            Label prefixHeader = new Label();
            Label numHeader = new Label();
            Label suffixHeader = new Label();
            Label empty = new Label();
            Label lengthOfNumHeader = new Label();

            prefixHeader.Size = new System.Drawing.Size(100, 30);
            numHeader.Size = new System.Drawing.Size(100, 30);
            suffixHeader.Size = new System.Drawing.Size(100, 30);
            empty.Size = new System.Drawing.Size(100, 30);
            lengthOfNumHeader.Size = new System.Drawing.Size(100, 30);

            prefixHeader.Text = "Префикс:";
            numHeader.Text = "Номер:";
            suffixHeader.Text = "Суффикс:";
            empty.Text = "   ";
            lengthOfNumHeader.Text = "Длина номера:";


            // Добавляем заголовки в обертку
            prefixHeader.Parent = headersWrapper;
            headersWrapper.Controls.Add(prefixHeader);
            prefixHeader.Anchor = AnchorStyles.Left;

            numHeader.Parent = headersWrapper;
            headersWrapper.Controls.Add(numHeader);
            numHeader.Anchor = AnchorStyles.Left;

            suffixHeader.Parent = headersWrapper;
            headersWrapper.Controls.Add(suffixHeader);
            suffixHeader.Anchor = AnchorStyles.Left;

            empty.Parent = headersWrapper;
            headersWrapper.Controls.Add(empty);
            empty.Anchor = AnchorStyles.Left;

            lengthOfNumHeader.Parent = headersWrapper;
            headersWrapper.Controls.Add(lengthOfNumHeader);
            lengthOfNumHeader.Anchor = AnchorStyles.Left;

            // Инициализируем обертку полей
            FlowLayoutPanel fieldsWrapper = new FlowLayoutPanel();
            fieldsWrapper.FlowDirection = FlowDirection.LeftToRight;
            fieldsWrapper.AutoSize = true;

            // Инициализируем поля
            prefixField.Size = new System.Drawing.Size(100, 30);
            numField.Size = new System.Drawing.Size(100, 30);
            suffixField.Size = new System.Drawing.Size(100, 30);
            lengthOfNumField.Size = new System.Drawing.Size(100, 30);


            Label newEmpty = new Label();
            newEmpty.Size = new System.Drawing.Size(100, 30);
            newEmpty.Text = "   ";
            newEmpty.Padding = new Padding(10, 0, 0, 0);

            numField.KeyPress += Num_KeyPress;
            lengthOfNumField.KeyPress += Num_KeyPress;

            // Добавляем поля в обертку
            prefixField.Parent = fieldsWrapper;
            fieldsWrapper.Controls.Add(prefixField);
            prefixField.Anchor = AnchorStyles.Left;

            numField.Parent = fieldsWrapper;
            fieldsWrapper.Controls.Add(numField);
            numField.Anchor = AnchorStyles.Left;

            suffixField.Parent = fieldsWrapper;
            fieldsWrapper.Controls.Add(suffixField);
            suffixField.Anchor = AnchorStyles.Left;

            newEmpty.Parent = fieldsWrapper;
            fieldsWrapper.Controls.Add(newEmpty);
            newEmpty.Anchor = AnchorStyles.Left;

            lengthOfNumField.Parent = fieldsWrapper;
            fieldsWrapper.Controls.Add(lengthOfNumField);
            lengthOfNumField.Anchor = AnchorStyles.Left;

            // Инициализируем кнопку
            Button button = new Button();
            button.Size = new Size(200, 60);
            button.Text = "Перенумеровать";
            button.Click += RunProgram;

            // Добавляем элементы в форму
            buildingsHeader.Parent = formWrapper;
            formWrapper.Controls.Add(buildingsHeader);
            buildingsHeader.Anchor = AnchorStyles.Left;

            buildingsCB.Parent = formWrapper;
            formWrapper.Controls.Add(buildingsCB);
            buildingsCB.Anchor = AnchorStyles.Left;

            tomesHeader.Parent = formWrapper;
            formWrapper.Controls.Add(tomesHeader);
            tomesHeader.Anchor = AnchorStyles.Left;

            tomesCB.Parent = formWrapper;
            formWrapper.Controls.Add(tomesCB);
            tomesCB.Anchor = AnchorStyles.Left;

            sheetHeader.Parent = formWrapper;
            formWrapper.Controls.Add(sheetHeader);
            sheetHeader.Anchor = AnchorStyles.Left;

            sheetsCB.Parent = formWrapper;
            formWrapper.Controls.Add(sheetsCB);
            sheetsCB.Anchor = AnchorStyles.Left;

            headersWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(headersWrapper);
            headersWrapper.Anchor = AnchorStyles.Left;

            fieldsWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(fieldsWrapper);
            fieldsWrapper.Anchor = AnchorStyles.Left;

            button.Parent = formWrapper;
            formWrapper.Controls.Add(button);
            button.Anchor = AnchorStyles.Top;
        }

        // Методы

        public void RepopulateComboBox(System.Windows.Forms.ComboBox comboBox, List<ViewSheet> list) 
        {
            comboBox.Items.Clear();
            bool flag = true;
            foreach (ViewSheet sheet in list)
            {
                comboBox.Items.Add(sheet.SheetNumber + "  |  " + sheet.Name);
                if (flag)
                {
                    FillTextBoxFields(sheet);
                    flag = false;
                }
            }
            comboBox.SelectedIndex = 0;            
        }
        public void RepopulateComboBox(System.Windows.Forms.ComboBox comboBox, Dictionary<string, List<ViewSheet>> dictionary)
        {
            comboBox.Items.Clear();
            foreach (var key in dictionary.Keys)
            {
                comboBox.Items.Add(key);
            }
            comboBox.SelectedIndex = 0;
        }

        public void FillTextBoxFields (ViewSheet sheet)
        {
            prefixField.Text = "";
            suffixField.Text = "";
            numField.Text = Regex.Replace(sheet.SheetNumber.ToString(), @"\p{C}+", string.Empty);
            lengthOfNumField.Text = numField.Text.Length.ToString();
        }

        // События

        public void OnBuildingChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.ComboBox comboBox = (System.Windows.Forms.ComboBox)sender;
            string newBKey = comboBox.SelectedItem.ToString();
            RepopulateComboBox(tomesCB, buildingsDict[newBKey]);
            string newTKey = tomesCB.SelectedItem.ToString();
            RepopulateComboBox(sheetsCB, buildingsDict[newBKey][newTKey]);
        }

        public void OnTomeChanged(object sender, EventArgs e)
        {            
            string newBKey = buildingsCB.SelectedItem.ToString();
            string newTKey = tomesCB.SelectedItem.ToString();
            RepopulateComboBox(sheetsCB, buildingsDict[newBKey][newTKey]);
        }

        public void OnSheetChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.ComboBox comboBox = (System.Windows.Forms.ComboBox)sender;
            int index = comboBox.SelectedIndex;
            string keyB = buildingsCB.SelectedItem.ToString();
            string keyT = tomesCB.SelectedItem.ToString();

            ViewSheet sheet = buildingsDict[keyB][keyT][index];
            FillTextBoxFields(sheet);
        }

        // Функция, запрещающая вводить в поля все, кроме цифр и точки
        private void Num_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar)) return;
            if (Char.IsControl(e.KeyChar)) return;
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.Contains('.') == false)) return;
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).SelectionLength == (sender as System.Windows.Forms.TextBox).TextLength)) return;
            e.Handled = true;
        }

        private void RunProgram (object sender, EventArgs e)
        {
            Transaction t = new Transaction(Doc, "Перенумеровать листы");
            t.Start();

            DockablePaneId dpId = DockablePanes.BuiltInDockablePanes.ProjectBrowser;
            DockablePane pB = new DockablePane(dpId);
            pB.Hide();
                   

            string keyB = buildingsCB.SelectedItem.ToString();
            string keyT = tomesCB.SelectedItem.ToString();
            int sheetIndex = sheetsCB.SelectedIndex;
            List<ViewSheet> sheetList = buildingsDict[keyB][keyT];

            ViewSheet previousSheet = sheetList[sheetIndex];
            ViewSheet currentSheet = sheetList[sheetIndex];

            List<string> oldNumbers = new List<string>();

            int zeroesPadding = Int32.Parse(lengthOfNumField.Text);

            // Убираем все перезаписываемые номера из списка номеров
            // Записываем старые номера в отдельный список
            // Перебиваем номера на временные, идиотские

            for (int i = sheetIndex; i < sheetList.Count; i++) 
            {
                sheetNums.Remove(sheetList[i].SheetNumber);
                oldNumbers.Add(sheetList[i].SheetNumber);
                sheetList[i].SheetNumber += "--Temp-Temp--" + i.ToString();
            }

            // Обрабатываем первый лист
            string tempNum = Regex.Replace(oldNumbers[0], @"\p{C}+", string.Empty); // Очищаем номер от лишних символов
            string[] prevSheet_old_numElements = tempNum.Split('.'); // Разделяем номер по точке
            string[] prevSheet_new_numElements = numField.Text.Split('.');
            prevSheet_new_numElements[0] = prevSheet_new_numElements[0].PadLeft(zeroesPadding, '0');

            string newNum = prefixField.Text + String.Join(".", prevSheet_new_numElements) + suffixField.Text;
            // Проверяем новый номер на уникальность
            while (!sheetNums.Add(newNum))
            {
                // Если есть знак после точки, меняем его, если нет, меняем основное число
                if (prevSheet_new_numElements.Count() > 1)
                {
                    int tempFollow = Int32.Parse(prevSheet_new_numElements[1]);
                    tempFollow++;
                    prevSheet_new_numElements[1] = tempFollow.ToString();
                }
                else
                {
                    int tempBase = Int32.Parse(prevSheet_new_numElements[0]);
                    tempBase++;
                    prevSheet_new_numElements[0] = tempBase.ToString().PadLeft(zeroesPadding, '0');
                }
                newNum = prefixField.Text + String.Join(".", prevSheet_new_numElements) + suffixField.Text;
            }

            previousSheet.SheetNumber = newNum;

            // Перебираем все остальные листы, сравнивая их с предыдущим
            for (int i = sheetIndex + 1; i < sheetList.Count; i++)
            {                
                currentSheet = sheetList[i];
                tempNum = Regex.Replace(oldNumbers[i - sheetIndex], @"\p{C}+", string.Empty);
                string[] currentSheet_old_numElements = tempNum.Split('.'); // Разделяем номер по точке
                List<string> currentSheet_new_numElements = new List<string>();

                if (prevSheet_old_numElements[0] == currentSheet_old_numElements[0])
                {
                    currentSheet_new_numElements.Add(prevSheet_new_numElements[0].PadLeft(zeroesPadding, '0')); // Если старые номера совпадают, присваиваем новый номер такой же, как у предыдущего листа
                    if (currentSheet_old_numElements.Count() > 1)                   // Если в текущем листе есть номер после точки...
                    {
                        if (prevSheet_new_numElements.Count() > 1)                  // ... и в предыдущем листе есть номер после точки...
                        {
                            int newFollow = Int32.Parse(prevSheet_new_numElements[1]);
                            newFollow++;
                            currentSheet_new_numElements.Add(newFollow.ToString());
                        }
                        else
                        {
                            currentSheet_new_numElements.Add("1");
                        }
                    }

                }
                else  // Если номера не совпадали, увеличиваем базу на 1
                {
                    int newBaseNum = Int32.Parse(prevSheet_new_numElements[0]);
                    newBaseNum++;
                    currentSheet_new_numElements.Add(newBaseNum.ToString().PadLeft(zeroesPadding, '0'));
                    if (currentSheet_old_numElements.Count() > 1)
                    {
                        currentSheet_new_numElements.Add("1");
                    }
                }

                // Проверяем новый номер на уникальность
                string[] currentSheetNewNumber = currentSheet_new_numElements.ToArray();
                newNum = prefixField.Text + String.Join(".", currentSheetNewNumber) + suffixField.Text;
                while (!sheetNums.Add(newNum)) 
                {
                    if (currentSheetNewNumber.Count() > 1)
                    {
                        int tempFollow = Int32.Parse(currentSheetNewNumber[1]);
                        tempFollow++;
                        currentSheetNewNumber[1] = tempFollow.ToString();
                    }
                    else                     
                    {
                        int tempBase = Int32.Parse(currentSheetNewNumber[0]);
                        tempBase++;
                        currentSheetNewNumber[0] = tempBase.ToString();
                    }
                    newNum = prefixField.Text + String.Join(".", currentSheetNewNumber) + suffixField.Text;
                }


                currentSheet.SheetNumber = newNum;

                previousSheet = currentSheet;
                prevSheet_new_numElements = currentSheetNewNumber;
                prevSheet_old_numElements = currentSheet_old_numElements;
            }

            pB.Show();
            t.Commit();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
