using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Manifold.Interop;
using Microsoft.VisualBasic;

namespace ExternalControlForum
{
    public class InternalScript
    {
        Document _doc;
        ComponentSet comps;

        public InternalScript(string MapFileLocation)
        {
            //On Initialisation, we want to be able to take in the document object, setting the _doc variable to this.
            Application app = new Application();
            Document doc = app.NewDocument(MapFileLocation, false);
            _doc = doc;
            comps = (ComponentSet)_doc.ComponentSet;
        }

        public void Run()
        {
       
            Query qry = _doc.NewQuery("Script Query", false);
            qry.Text = "SELECT TOP 1 * FROM [CONTROL TABLE] ORDER BY [DATESTAMP] ASC;";

            try
            {
                Console.WriteLine(DateTime.Now.ToString() + ": Running " + qry.Text);
                qry.RunEx(true);
            }
            catch (Exception)
            {
                _doc.Close(false);
                throw;
            }
        
            Table table = qry.Table;
          
            int varValue;
            DateTime varTime;


            if (this.TableHasRecords(table))
            {
                Record rec = table.RecordSet[0];
                varValue = (int)rec.get_Data("VALUE");
                varTime = (DateTime)rec.get_Data("DATESTAMP");


            }
            else
            {
                comps.Remove("Script Query");
                _doc.Close(false);
                throw new NullReferenceException("[CONTROL TABLE] must have records");
            }


            Console.WriteLine(DateTime.Now.ToString() + ": " + varValue.ToString() + " - " + varTime.ToString());

            qry.Text = "UPDATE [Drawing] SET [VALUE] = " + varValue +";";
            try
            {
                Console.WriteLine(DateTime.Now.ToString() + ": Running " + qry.Text);
                qry.RunEx(true);
            }
            catch (Exception)
            {
                comps.Remove("Script Query");
                _doc.Close(false);
                throw;
            }


            ExportPdf expt = (ExportPdf)_doc.NewExport("PDF");

            try 
	        {
                Console.WriteLine(DateTime.Now.ToString() + ": Exporting drawing to Layout " + @"C:\temp\Layout_" + varValue + ".pdf");
			    expt.Export(comps["Layout"], @"C:\temp\Layout_" + varValue + ".pdf");
	        }
	        catch (Exception)
	        {
                comps.Remove("Script Query");
                _doc.Close(false);
		        throw;
	        }
           

            qry.Text = "DELETE FROM [CONTROL TABLE] WHERE CStr([DATESTAMP]) = \"" + varTime + "\";";
            try
            {
                Console.WriteLine(DateTime.Now.ToString() + ": Running " + qry.Text);
                qry.RunEx(true);
            }
            catch (Exception)
            {
                comps.Remove("Script Query");
                _doc.Close();
                throw;
            }

            //final cleanup
            _doc.Close(false);
            _doc = null;
           
        }



        public bool TableHasRecords(Table table){
            int recCount = table.RecordSet.Count;
            if (recCount > 0)
                return true;
            else
                return false;
        }

        public void InsertDummyData()
        {


            Query qry = _doc.NewQuery("Ext Query", true);
            qry.Text = "INSERT INTO [CONTROL TABLE] ([VALUE], [DATESTAMP]) VALUES(((Rnd*100)*Rnd)*10, NOW);";
            qry.RunEx(true);
            comps.Remove("Ext Query");
            

        }










//msgbox varValue&" - "&CStr(varTime)

//qry.Text = "UPDATE [Drawing] SET [VALUE] = "&varValue&";"
//qry.RunEx True

//Set expt = Application.NewExport("PDF")

//' -- THIS LINE CAUSES THE SCRIPT TO FAIL
//' -- expt.Export comps("Layout"), "C:\Layout_"&varValue&".pdf"

//qry.Text = "DELETE FROM [CONTROL TABLE] WHERE CStr([DATESTAMP]) = "&Chr(34)&varTime&Chr(34)&";"
//qry.RunEx True

//comps.Remove("Script Query")

//End Sub
    }
}
