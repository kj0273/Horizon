using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ConsoleApp1
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application objApplication = null;
            SolidEdgeDraft.DraftDocument objDraftDocument = null;

            try
            {
                OleMessageFilter.Register();

                // Create/get the application with specific settings
                objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                // connect to current draftdocument
                objDraftDocument = (SolidEdgeDraft.DraftDocument)objApplication.ActiveDocument;

                //get the path and dft name
                string path = "" + objDraftDocument.FullName;

                //wipe out ".dft"
                path = path.Substring(0, path.Length - 4);

                //Saves the draft document as pdf and dwg
                objDraftDocument.SaveAs(path+".pdf");
                objDraftDocument.SaveAs(path + ".dwg");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }
    }
}
