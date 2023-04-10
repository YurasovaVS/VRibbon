using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace DeleteSignatures
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteSignaturesDWG : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            using (System.Windows.Forms.Form form = new SKRibbon.Forms.DeleteSigForm(doc))
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

            /*
            ICollection<Element> signatures = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).ToElements();
            StringBuilder sb = new StringBuilder();


            Transaction t = new Transaction(doc, "Удалить подписи");
            t.Start();
            foreach (ImportInstance signature in signatures)
            {                
                string name = signature.LookupParameter("Имя").AsString();
                sb.AppendLine(name);
                string pattern = "подпись_*";
                if (Regex.Match(name, pattern).Success)
                {
                    sb.AppendLine("Удалена");
                    doc.Delete(signature.Id);
                } 
            }
            TaskDialog.Show("Результат", sb.ToString());
            t.Commit();
            */
            return Result.Succeeded;  
        }
    }
}
