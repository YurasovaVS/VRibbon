using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using FakeArea;

namespace AddSignatures
{
    [Transaction(TransactionMode.Manual)]
    public class AddSignaturesDWG : IExternalCommand
    {
        public struct Signature 
        {
            public string path;
            public ImportInstance sigInstance;

            public Signature(string path, ImportInstance instance)
            {
                this.path = path;
                this.sigInstance = instance;
            }
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Берем документ
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            //Прописываем путь к подписям
            string path = "\\\\ABSKNAS\\переезд\\13_Пользователи\\_ПодписиСК\\";

            using (System.Windows.Forms.Form form = new SKRibbon.AddSigForm(doc, path))
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }
        }
    }
}
