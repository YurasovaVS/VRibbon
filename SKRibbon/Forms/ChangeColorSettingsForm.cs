using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Creation;
using System.Windows.Forms;
using ColorizeTabs;
using Autodesk.Revit.UI;

namespace SKRibbon
{
    public partial class ChangeColorSettingsForm : Form
    {
        FlowLayoutPanel formWrapper = new FlowLayoutPanel();
        FlowLayoutPanel colorsWrapper = new FlowLayoutPanel();
        UIApplication UiApp;
        public ChangeColorSettingsForm(UIApplication uiApp)
        {
            InitializeComponent();

            UiApp = uiApp;

            // Инициализация формвраппера
            formWrapper.FlowDirection = FlowDirection.TopDown;
            formWrapper.AutoSize = true;

            // Инициализация цветов
            colorsWrapper.FlowDirection = FlowDirection.TopDown;
            colorsWrapper.AutoSize = true;

            string[] colorHexes= Properties.appSettings.Default.tabColors.Split(',');
            int i = 1;
            foreach (string hex in colorHexes) {
                AddColorRow(i.ToString(), hex);
                i++;
            }

            // Инициализация кнопок
            FlowLayoutPanel buttonWrapper = new FlowLayoutPanel();
            buttonWrapper.FlowDirection = FlowDirection.LeftToRight;
            buttonWrapper.AutoSize = true;

            Button plusButton = new Button();
            Button minusButton = new Button();
            Button okButton = new Button();

            plusButton.Text = "+";
            minusButton.Text = "-";
            okButton.Text = "OK";

            plusButton.Size = minusButton.Size = okButton.Size = new Size(30, 30);
            plusButton.Margin = minusButton.Margin = new Padding (0, 5, 5, 5);
            okButton.Margin = new Padding (130, 5, 0, 5);

            plusButton.Anchor = minusButton.Anchor = AnchorStyles.Left;
            okButton.Anchor = AnchorStyles.Right;

            plusButton.Click += AddColor;
            minusButton.Click += RemoveColor;
            okButton.Click += SaveNewSettings;

            plusButton.Parent = minusButton.Parent = okButton.Parent = buttonWrapper;
            buttonWrapper.Controls.Add(plusButton);
            buttonWrapper.Controls.Add(minusButton);
            buttonWrapper.Controls.Add(okButton);

            // Добавление элементов во formwrapper
            colorsWrapper.Parent = buttonWrapper.Parent = formWrapper;
            formWrapper.Controls.Add(colorsWrapper);
            formWrapper.Controls.Add(buttonWrapper);

            formWrapper.Parent = this;
            this.Controls.Add(formWrapper);

            this.Width = 290;
            this.Height = 300;
            this.AutoScroll = true;
            this.MinimumSize = new Size (this.Width, this.Height);
            this.MaximumSize = new Size (this.Width, 600);
        }

        private void AddColorRow (string num, string hex)
        {
            // Wrapper
            // ( [Num] [Color] [Hex] )

            int height = 25;

            FlowLayoutPanel wrapper = new FlowLayoutPanel();
            wrapper.FlowDirection = FlowDirection.LeftToRight;
            wrapper.AutoSize = true;
            wrapper.Padding = new Padding(2, 2, 2, 2);
            wrapper.Anchor = AnchorStyles.Top;

            Label numLabel = new Label();
            numLabel.Size = new Size(20, height);
            numLabel.Text = num;
            numLabel.Anchor = AnchorStyles.Left;

            Button colorButton = new Button();
            colorButton.Size = new Size(150, height);
            colorButton.Text = "";
            colorButton.BackColor = ColorTranslator.FromHtml(hex);
            colorButton.Anchor = AnchorStyles.Left;
            colorButton.Click += ChooseColor;

            Label hexLabel = new Label();
            hexLabel.Size = new Size(70, height);
            hexLabel.Text = hex;
            hexLabel.Anchor = AnchorStyles.Left;

            // Добавляем элементы во враппер

            numLabel.Parent = wrapper;
            colorButton.Parent = wrapper;
            hexLabel.Parent = wrapper;

            wrapper.Controls.Add(numLabel);
            wrapper.Controls.Add(colorButton);
            wrapper.Controls.Add(hexLabel);

            wrapper.Parent = colorsWrapper;
            colorsWrapper.Controls.Add(wrapper);                        
        }

        private void ChooseColor(object sender, System.EventArgs e)
        {
            Button button = sender as Button;
            FlowLayoutPanel wrapper = (FlowLayoutPanel)button.Parent;
            Label hexLabel = (Label)wrapper.Controls[2];

            ColorDialog MyDialog = new ColorDialog();

            // Update the text box color if the user clicks OK 
            if (MyDialog.ShowDialog() == DialogResult.OK)
            {
                button.BackColor = MyDialog.Color;
                hexLabel.Text = ColorTranslator.ToHtml(MyDialog.Color);
            }                
        }

        private void AddColor(object sender, System.EventArgs e)
        {
            AddColorRow((colorsWrapper.Controls.Count + 1).ToString(), "#AABBCC");
        }

        private void RemoveColor(object sender, System.EventArgs e)
        {
            int i = colorsWrapper.Controls.Count - 1;
            FlowLayoutPanel row = (FlowLayoutPanel)colorsWrapper.Controls[i];
            colorsWrapper.Controls.Remove(row);
        }

        private void SaveNewSettings(object sender, System.EventArgs e)
        {
            string newSettings = "";
            foreach (FlowLayoutPanel panel in colorsWrapper.Controls)
            {
                Label hex = (Label)panel.Controls[2];
                newSettings += hex.Text + ",";
            }
            Properties.appSettings.Default.tabColors = newSettings.Substring(0, newSettings.Length - 1);
            Properties.appSettings.Default.Save();
            ColorizeTabs.ColorizeTabs.RunCommand(UiApp, Properties.appSettings.Default.tabColorFlag);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
