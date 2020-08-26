using SolidEdgeCommunity;
using System;
using System.Runtime.InteropServices;

namespace Easy_Save_as
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgeDraft.DraftDocument objDraftDocument = null;

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();

                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                //check the current file belongs to par,asm or dft file type
                if (application.ActiveDocumentType.ToString().Equals("igPartDocument"))
                {
                    partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;
                    //if there is a par exist
                    if (partDocument != null)
                    {
                        //get the path and file name
                        string path = "" + partDocument.FullName;

                        //wipe out file type name
                        path = path.Substring(0, path.Length - 4);

                        //save as .stp and .igs
                        partDocument.SaveCopyAs(path + ".stp");
                        partDocument.SaveCopyAs(path + ".igs");
                    }
                    else
                    {
                        Console.WriteLine("Can't find the file, press anything to continue");
                        Console.ReadLine();
                    }
                }
                else if (application.ActiveDocumentType.ToString().Equals("igAssemblyDocument"))
                {
                    assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument;
                    //if there is an asm exist
                    if (assemblyDocument != null)
                    {
                        //get the path and file name
                        string path = "" + assemblyDocument.FullName;

                        //wipe out file type name
                        path = path.Substring(0, path.Length - 4);

                        //save as .stp and .igs
                        assemblyDocument.SaveCopyAs(path + ".stp");
                        assemblyDocument.SaveCopyAs(path + ".igs");
                    }
                    else
                    {
                        Console.WriteLine("Can't find the file, press anything to continue");
                        Console.ReadLine();
                    }
                }
                else if(application.ActiveDocumentType.ToString().Equals("igDraftDocument"))
                {
                    objDraftDocument = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;
                    if(objDraftDocument != null)
                    {
                        //get the path and dft name
                        string path = "" + objDraftDocument.FullName;

                        //wipe out ".dft"
                        path = path.Substring(0, path.Length - 4);

                        //Saves the draft document as pdf and dwg
                        objDraftDocument.SaveAs(path + ".pdf");
                        objDraftDocument.SaveAs(path + ".dwg");
                    }
                    else
                    {
                        Console.WriteLine("Can't find the file, press anything to continue");
                        Console.ReadLine();
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }
    }
}
