using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SKRibbon
{
    public class FormDesign
    {
        // Цвет фона форм
        Color FormBcgrColor = Color.White;
        Color ElementsBcgrColor = Color.LavenderBlush;
        // Текстбокс
        public class VTextBox : TextBox
        { 
            public VTextBox()
            {
                BorderStyle = BorderStyle.FixedSingle;                
                BackColor = Color.FromArgb(255, 251, 245, 255);
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
                ForeColor = GeneralColor;
                Font = new Font(Button.DefaultFont, FontStyle.Bold);
            }
        }

        public class VHeaderLabel : Label
        {
            Color GeneralColor = Color.FromArgb(255, 174, 112, 199);
            public VHeaderLabel()
            {
                Font = new Font("Arial", 12, FontStyle.Bold);
                ForeColor = GeneralColor;
            }

        }
    }
}
