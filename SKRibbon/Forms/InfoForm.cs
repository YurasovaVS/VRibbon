using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using Autodesk.Revit.DB;


namespace SKRibbon.Forms
{
    public partial class InfoForm : System.Windows.Forms.Form
    {
        Document Doc;
        public InfoForm(Document doc)
        {
            /*
             * 
             * 
            Программа предназначена для автоматизации рутинных задач и сокращения времени на них.
            Пока она работает, вы можете попить кофе и посплетничать с коллегами :)
                        
            
            Исключительное право на данный продукт принадлежит разработчику.
            Согласно желанию правообладателя, данный продукт является бесплатным и распространяется по лицензии GNU (GPL).
            Продажа или взимание платы за данный продукт является незаконной и попадает под Статью 159 УК РФ.
            Если 

            
             */
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Form form = new LicenseForm();
            form.ShowDialog();
            this.Close();
            this.DialogResult = DialogResult.OK;
        }
    }
}
