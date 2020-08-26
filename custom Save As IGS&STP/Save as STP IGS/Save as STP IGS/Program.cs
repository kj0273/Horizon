using SolidEdgeCommunity;
using System;
using System.Runtime.InteropServices;

namespace Save_as_STP_IGS
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
            SolidEdgePart.PartDocument partDocument = null;

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();

                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                //check the current file belongs to par or asm file type
                if (application.ActiveDocumentType.ToString().Equals("igPartDocument"))
                {
                    partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;
                }
                if (application.ActiveDocumentType.ToString().Equals("igAssemblyDocument"))
                {
                    assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument;
                }


                //if there is an asm exist
                if (assemblyDocument != null)
                {
                    //get the path and file name
                    string path = "" + assemblyDocument.FullName;

                    //wipe out file type name
                    path = path.Substring(0, path.Length - 4);

                    assemblyDocument.SaveCopyAs(path + ".stp");
                    assemblyDocument.SaveCopyAs(path + ".igs");
                }

                //if there is a par exist
                if (partDocument != null)
                {
                    //get the path and file name
                    string path = "" + partDocument.FullName;

                    //wipe out file type name
                    path = path.Substring(0, path.Length - 4);

                    partDocument.SaveCopyAs(path + ".stp");
                    partDocument.SaveCopyAs(path + ".igs");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }
    }
}