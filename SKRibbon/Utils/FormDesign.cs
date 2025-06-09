using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static SKRibbon.FormDesign;

namespace SKRibbon
{
    public class FormDesign
    {
        // Цвет фона форм
        Color FormBcgrColor = Color.White;
        Color ElementsBcgrColor = Color.LavenderBlush;

        // Форма
        public class VForm : Form
        {
            public VForm()
            {
                AutoSize = true;
                AutoScroll = true;
                ShowIcon = false;
                FormBorderStyle = FormBorderStyle.FixedSingle;
                BackColor = System.Drawing.Color.White;
                Text = "";
            }
        }
        // Текстбокс
        public class VTextBox : TextBox
        { 
            public VTextBox()
            {
                BorderStyle = BorderStyle.FixedSingle;                
                BackColor = Color.White;
            }
        }

        public class VButton : Button
        {
            Color GeneralColor = Color.FromArgb(255, 174, 112, 199);
            public VButton()
            {
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 3;
                FlatAppearance.BorderColor = GeneralColor;
                FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 251, 245, 255);
                FlatAppearance.MouseDownBackColor = Color.FromArgb(255, 174, 112, 199);
                ForeColor = GeneralColor;
                Font = new Font(Button.DefaultFont, FontStyle.Bold);
            }
        }

        public class VHeaderLabel : Label
        {
            Color GeneralColor = Color.FromArgb(255, 174, 112, 199);
            public VHeaderLabel()
            {

                Padding = new Padding(10, 5, 0, 10);
                Font = new Font("Arial", 12, FontStyle.Bold);
                ForeColor = GeneralColor;
            }
        }


        // Неиспользуемые классы
        public class VHeader : FlowLayoutPanel
        {
            public VHeader(string text, int width)
            {
                Height = 60;
                Width = width;
                FlowDirection = FlowDirection.LeftToRight;
                BackColor = Color.FromArgb(255, 174, 112, 199);

                Label header = new Label();
                header.Text = text;
                header.Height = 40;
                if (width > 50) header.Width = width - 50;
                header.ForeColor = Color.White;
                header.Font = new Font("Arial", 12, FontStyle.Bold);
                header.Anchor = AnchorStyles.Left;
                header.Padding = new Padding(15, 15, 0, 0);

                header.Parent = this;
                this.Controls.Add(header);

                Button closeButton = new Button();
                closeButton.Text = "\u271A";
                closeButton.Size = new Size(30, 30);
                closeButton.FlatStyle = FlatStyle.Flat;
                closeButton.FlatAppearance.BorderSize = 1;
                closeButton.FlatAppearance.BorderColor = Color.White;

                closeButton.Anchor = AnchorStyles.Left;
                closeButton.Padding = new Padding(0, 15, 0, 0);

                closeButton.Parent = this;
                this.Controls.Add(closeButton);
            }
        }
    }
}
