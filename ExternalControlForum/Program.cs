using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExternalControlForum
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Manifold console demonstration.  Please press any key to continue");
            Console.ReadLine();
            InternalScript script = new InternalScript(@"c:\temp\EXTERNAL CONTROL FORUM.map");
            //Create the insert query in the manifold file, and inserts the dummy data
            script.InsertDummyData();
            //runs the internal script (although now it is in the Internal Script class rather than within the 
            //project.
            script.Run();
            Console.WriteLine("Process complete.");
            Console.ReadLine();

        }
    }
}
