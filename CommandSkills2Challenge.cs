#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;


#endregion

namespace RAB_Skills02
{
    [Transaction(TransactionMode.Manual)]
    public class CommandSkills2Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                //create levelsList from a csv file using OpenFileDialog
                using (OpenFileDialog openfiledialog = new OpenFileDialog())
                {
                    openfiledialog.InitialDirectory = "c : \\";
                    openfiledialog.Filter = "csv files (*.csv)|*.csv|allfiles(*.*)|*.*";
                    //openfiledialog.FilterIndex = 1;
                    //openfiledialog.RestoreDirectory = true;

                    if (openfiledialog.ShowDialog() == DialogResult.OK)
                    {
                        // get the csv file for the levelsList
                        string levfilepath = openfiledialog.FileName;

                        //create an array from the levelsList form a csv file
                        string[] levArraydata = File.ReadAllLines(levfilepath);

                        //Remove the header row
                        //arraydata.Skip(0).ToArray();
                        levArraydata = levArraydata.Where(w => w != levArraydata[0]).ToArray();
                        //create a transaction for the mylevelsList

                        Transaction levtransaction = new Transaction(doc);
                        levtransaction.Start("create levelsList");

                        foreach (string item in levArraydata)
                        {
                            //Create an instance of the "MyLevelStruct" structure adn populate the data from the levelsList csv file                            
                            MylevelsStruct mylevStruct = new MylevelsStruct();
                            mylevStruct.Name = item.Split(',')[0];
                            mylevStruct.Value1 = item.Split(',')[1];
                            mylevStruct.Value2 = item.Split(',')[2];

                            List<MylevelsStruct> mylevList = new List<MylevelsStruct>();
                            mylevList.Add(mylevStruct);

                            //loop through the data (mySheetsList) 
                            foreach (MylevelsStruct level in mylevList)
                            {
                                //change level elevation values from string to double
                                double levmvalue = Convert.ToDouble(level.Value2);

                                //change level elevation values unit from meters to feet
                                double levfvalue = UnitUtils.Convert(levmvalue, UnitTypeId.Meters, UnitTypeId.Feet);

                                //create and name the levelsList
                                Level lev = Level.Create(doc, levfvalue);
                                lev.Name = level.Name;
                                //elementid levelid = lev.id;
                            }
                        }
                        //commit the levelsList transaction
                        levtransaction.Commit();

                    }
                }
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }

            try
            {
                //Create Viewplanes for each level in the doc

                //Get viewFamilyTypes
                FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
                vftCollector.OfClass(typeof(ViewFamilyType));

                //Get levelsList
                FilteredElementCollector levelsCollector = new FilteredElementCollector(doc);
                levelsCollector.OfClass(typeof(Level))
                    .WhereElementIsNotElementType();

                //Create a transaction for the viewplane creation
                Transaction viewPlnTransaction = new Transaction(doc);
                viewPlnTransaction.Start("Create viewPlans");

                foreach (Level lvl in levelsCollector)
                {
                    ViewFamilyType planVFT = null;
                    ViewFamilyType rcpVFT = null;

                    foreach (ViewFamilyType vft in vftCollector)
                    {

                        if (vft.ViewFamily == ViewFamily.FloorPlan)
                            planVFT = vft;

                        if (vft.ViewFamily == ViewFamily.CeilingPlan)
                            rcpVFT = vft;
                    }

                    //Create plan and RCP views
                    ViewPlan floorlan = ViewPlan.Create(doc, planVFT.Id, lvl.Id);
                    ViewPlan rcpPlan = ViewPlan.Create(doc, rcpVFT.Id, lvl.Id);

                    //Rename the RCP plans to add "RCP" suffix
                    rcpPlan.Name = rcpPlan.Name + " RCP";

                }
                //Commit and dispose the viewPlans transaction

                viewPlnTransaction.Commit();

            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }


            //Create sheets from a csv file using OpenFileDialog

            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C : \\";
                    openFileDialog.Filter = "csv files(*.csv)|*.csv|All files(*.*)|*.*";
                    //openFileDialog.FilterIndex = 2;
                    //openFileDialog.RestoreDirectory = true;



                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Create the active document sheets
                        // get the CSV file for the sheets
                        string ShtsfilePath = openFileDialog.FileName;

                        //Read the CSV lines
                        string[] sheetsArrayData = File.ReadAllLines(ShtsfilePath);

                        //Remove the header row
                        //sheetsArrayData.Skip(0).ToArray();
                        sheetsArrayData = sheetsArrayData.Where(w => w != sheetsArrayData[0]).ToArray();

                        //Create a transaction for the sheets
                        Transaction sheetsTransaction = new Transaction(doc);
                        sheetsTransaction.Start("Create Sheets");

                        foreach (var item in sheetsArrayData)
                        {
                            //Create an instance of the "mySheetsStruct" structure and populate the data from the SheetsList csv file  
                            MySheetsStruct mySheetsStruct = new MySheetsStruct();
                            mySheetsStruct.Name = item.Split(',')[0];
                            mySheetsStruct.Number = item.Split(',')[1];

                            //Create a list of mySheetsStruct
                            List<MySheetsStruct> mySheetsList = new List<MySheetsStruct>();
                            mySheetsList.Add(mySheetsStruct);

                            //Using the GetTitleBlockByName method (with a ENU template Imperial-Architectural Template")
                            Element tb = GetTitleBlockByName(doc, "E1 30x42 Horizontal");
                            //Using the GetTitleBlockByName method (with a french template "Gabarit de Construction")
                            //Element tb = GetTitleBlockByName(doc, "Métrique A1");
                            ElementId tbId = tb.Id;

                            //Loop through the data (Sheets)
                            foreach (MySheetsStruct sheet in mySheetsList)
                            {
                                //Create, number and name the sheets
                                ViewSheet vSheet = ViewSheet.Create(doc, tbId);
                                vSheet.Name = sheet.Number;
                                vSheet.SheetNumber = sheet.Name;
                            }
                        }
                        //Commit the sheets Transaction
                        sheetsTransaction.Commit();
                        sheetsTransaction.Dispose();

                    }
                }
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }

            //insert views in sheets

            try
            {
                //Create a transaction for the Insertion of views in sheets
                Transaction trans = new Transaction(doc);
                trans.Start("InsertViews in sheets");

                // Collect sheets in the active document
                IList<ViewSheet> viewSheets = new FilteredElementCollector(doc)
                     .OfCategory(BuiltInCategory.OST_Sheets)
                     .WhereElementIsNotElementType()
                     .Cast<ViewSheet>()
                     .ToList();


                FilteredElementCollector vpCollector = new FilteredElementCollector(doc);
                IList<Autodesk.Revit.DB.View> views = vpCollector.OfClass(typeof(ViewPlan))
                    .WhereElementIsNotElementType()
                    .Cast<Autodesk.Revit.DB.View>()
                    .ToList();

                //testing the geViewByName without using it 
                //Element view = GetViewByName(doc, "Level01");

                foreach (var vSheet in viewSheets)
                {
                    foreach (var v in views)
                    {
                        if (v != null && !v.IsTemplate)
                        {
                            // Get the middle point of the sheet (insertion point)
                            BoundingBoxUV outline = vSheet.Outline;
                            double x = (outline.Max.U + outline.Min.U) / 2;
                            double y = (outline.Max.V + outline.Min.V) / 2;

                            XYZ midpt = new XYZ(x, y, 0);

                            //Create the viewport
                            if (Viewport.CanAddViewToSheet(doc, vSheet.Id, v.Id) && (vSheet.Name.Contains(v.Name)))
                            {
                                Viewport vport = Viewport.Create(doc, vSheet.Id, v.Id, midpt);
                            }
                        }
                    }
                }
                trans.Commit();

            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
            return Result.Succeeded;

        }

        internal Element GetTitleBlockByName(Document doc, string name)
        {
            FilteredElementCollector tbCollector = new FilteredElementCollector(doc);
            tbCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            foreach (Element tb in tbCollector)
            {

                if (tb.Name == name)

                    return tb;
            }

            return null;

        }

        internal Element GetViewByName(Document doc, string name)
        {
            FilteredElementCollector vCollecor = new FilteredElementCollector(doc);
            vCollecor.OfCategory(BuiltInCategory.OST_Views);

            foreach (Element v in vCollecor)
            {

                if (v.Name == name)
                    return v;
            }
            return null;
        }

        struct MylevelsStruct
        {
            public string Name;
            public string Value1;
            public string Value2;
            public MylevelsStruct(string name, string value1, string value2)
            {
                Name = name;
                Value1 = value1;
                Value2 = value2;
            }
        }

        struct MySheetsStruct
        {
            public string Name;
            public string Number;

            public MySheetsStruct(string name, string number)
            {
                Name = name;
                Number = number;
            }
        }
    }
}
