using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeCommunity;

namespace program2
{
    class Program
    {
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application setApplication = null;
            SolidEdgeFramework.Documents setDocuments = null;
            SolidEdgePart.PartDocument setPartDocument = null;

            try
            {
                //the task start here
                setApplication = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);

                //obtain the reference with Documents
                setDocuments = setApplication.Documents;

                //create a new document
                setPartDocument = (PartDocument)setDocuments.Add("SolidEdge.PartDocument");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                //insert the obj in reverse
                setPartDocument = null;
                setDocuments = null;
                setApplication = null;

            }
        }
    }
}
