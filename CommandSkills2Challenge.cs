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
                //create levels from a csv file using OpenFileDialog
                using (OpenFileDialog openfiledialog = new OpenFileDialog())
                {
                    openfiledialog.InitialDirectory = "c : \\";
                    openfiledialog.Filter = "csv files (*.csv)|*.csv|allfiles(*.*)|*.*";
                    //openfiledialog.FilterIndex = 1;
                    //openfiledialog.RestoreDirectory = true;

                    if (openfiledialog.ShowDialog() == DialogResult.OK)
                    {
                        // get the csv file for the levels
                        string levfilepath = openfiledialog.FileName;

                        //create the active document levels form a csv file
                        string[] arraydata = File.ReadAllLines(levfilepath);

                        //create a list for csv lines
                        List<string> levels = new List<string>();
                        levels.AddRange(arraydata);

                        //remove the header row
                        levels.RemoveAt(0);

                        //create a transaction for the levels
                        Transaction levtransaction = new Transaction(doc);
                        levtransaction.Start("create levels");

                        //loop through the data (levels) 
                        foreach (var level in levels)
                        {
                            //use string.split method to separate text file data
                            string levelname = level.Split(',')[0];
                            string levelvalue = level.Split(',')[2];

                            //change level elevation values from string to double
                            double levmvalue = Convert.ToDouble(levelvalue);

                            //change level elevation values unit from meters to feet
                            double levfvalue = UnitUtils.Convert(levmvalue, UnitTypeId.Meters, UnitTypeId.Feet);

                            //create and name the levels
                            Level levv = Level.Create(doc, levfvalue);
                            levv.Name = levelname;
                            //elementid levelid = levv.id;

                        }
                        //commit the levels transaction
                        levtransaction.Commit();
                        levtransaction.Dispose();
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

                //Create a transaction for the viewplane creation
                Transaction viewPlnTransaction = new Transaction(doc);
                viewPlnTransaction.Start("Create viewPlans");

                //Get viewFamilyTypes
                FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
                vftCollector.OfClass(typeof(ViewFamilyType));

                //Get levels
                FilteredElementCollector levelsCollector = new FilteredElementCollector(doc);
                levelsCollector.OfClass(typeof(Level))
                    .WhereElementIsNotElementType();


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
                viewPlnTransaction.Dispose();


                //Create sheets from a csv file using OpenFileDialog
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

                        //Create a list for csv lines
                        List<string> sheets = new List<string>();
                        sheets.AddRange(sheetsArrayData);

                        //Remove the header row
                        sheets.RemoveAt(0);

                        //Create a transaction for the sheets
                        Transaction sheetsTransaction = new Transaction(doc);
                        sheetsTransaction.Start("Create Sheets");

                        //Loop through the data (Sheets)
                        foreach (var sheet in sheets)
                        {
                            //Create the vftCollector and get the title block element ID
                            FilteredElementCollector tBlockcollector = new FilteredElementCollector(doc);
                            ElementId titleBlockTypeId = tBlockcollector.OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElementId();

                            //Use String.split method to separate text file data
                            string sheetNumber = sheet.Split(',')[0];
                            string sheetName = sheet.Split(',')[1];

                            //Create, number and name the sheets
                            ViewSheet vSheet = ViewSheet.Create(doc, titleBlockTypeId);
                            vSheet.Name = sheetName;
                            vSheet.SheetNumber = sheetNumber;
                        }
                        //Commit the sheets Transaction
                        sheetsTransaction.Commit();
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

                //FilteredElementCollector shCollector = new FilteredElementCollector(doc);
                //shCollector.OfClass(typeof(ViewSheet));

                FilteredElementCollector vpCollector = new FilteredElementCollector(doc);
                IList<Autodesk.Revit.DB.View> views = vpCollector.OfClass(typeof(ViewPlan))
                    .WhereElementIsNotElementType()
                    .Cast<Autodesk.Revit.DB.View>()
                    .ToList();

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
                trans.Dispose();
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
