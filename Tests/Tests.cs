using System;
using System.Collections.Generic;
using NUnit.Framework;
using Shouldly;
using System.Text;
using Manifold.Interop;
using ExternalControlForum;


namespace Tests
{
    public class Tests
    {
        Application app;
        Document doc;

        string MapFilePath = @"c:\temp\EXTERNAL CONTROL FORUM.map";

        [Test]
        public void FirstTest()
        {
            true.ShouldBe(true); 
        }


        [SetUp]//Runs before every test in the class.
        public void Init()
        {
            //Create the app and doc objects, using an in memory map file for testing purposes.
            app = new Application();
            doc = app.NewDocument(MapFilePath, false);

        }

        [TearDown]//To be run after each test is finished.
        public void Cleanup()
        {
            //cleanup the doc and app objects.
            doc.Close(false);
            app.Quit();
            doc = null;
        }

        [Test]
        public void InternalScript_TableResult_Should_ReturnFalse_IfThereAreNoRecords()
        {
            //Create a new table with no records (or columns)
           
            Table testTable = (Table)doc.NewTable("test table",Type.Missing,true);
            InternalScript script = new InternalScript(MapFilePath);
            script.TableHasRecords(testTable).ShouldBe(false);
            
        }

        [Test]
        public void InternalScript_TableResult_Should_ReturnTrue_IfThereAreRecords()
        {
            //Create a new table with no records (or columns)

            Table testTable = (Table)doc.NewTable("test table", Type.Missing, true);
            Column col = testTable.ColumnSet.NewColumn();
            col.Name = "OID";
            testTable.ColumnSet.Add(col);
            Query q = doc.NewQuery("q");
            q.Text = "Insert into [test table](OID) values (1)";
            q.RunEx(true);
            InternalScript script = new InternalScript(MapFilePath);
            script.TableHasRecords(testTable).ShouldBe(true);

        }

        [Test]
        public void ScriptRuns()
        {
            InternalScript script = new InternalScript(MapFilePath);
            script.InsertDummyData();
            script.Run();
        }

    }
}
