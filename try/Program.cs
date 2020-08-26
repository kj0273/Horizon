using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using SolidEdgeCommunity;
using SolidEdgeFileProperties;
using SolidEdgeSDK;


namespace @try
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgeDraft.DraftDocument draftDocument = null;

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();


                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;
                draftDocument = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;


                /*
                application = new DesignManager.Application();
                propertySets = (DesignManager.PropertySets)application.PropertySets;
                propertySets.Open(@"C:\Part1.par", false);
                properties = (DesignManager.Properties)propertySets.Item["Custom"];
                property = (DesignManager.Property)properties.Add("My Custom Property", "My Custom String");
                propertySets.Save();
                */


                if (draftDocument != null)
                {
                    var NewName = @"C:\Users\win10-20200715\Desktop\output\Test.dft";
                    //partDocument.SaveAs(NewName, null, true);
                    draftDocument.SaveAs(NewName,null,true);
                    //partDocument.Save();
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
        }
    }
}
