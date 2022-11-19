#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;




#endregion

namespace RAB_Skills02
{
    [Transaction(TransactionMode.Manual)]
    public class CommandSkills2Challenge : IExternalCommand
    {

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C : \\";
                    openFileDialog.Filter = "csv files (*.csv) | *.csv | All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Get the CSV file for the levels
                        string LevfilePath = openFileDialog.FileName;

                        //Create the active document levels form a CSV file
                        string[] arrayData = File.ReadAllLines(LevfilePath);

                        //Create a list for csv lines
                        List<string> levels = new List<string>();
                        levels.AddRange(arrayData);

                        //Remove the header row
                        levels.RemoveAt(0);

                        //Create a transaction for the levels
                        Transaction levTransaction = new Transaction(doc);
                        levTransaction.Start("Create levels");

                        //Loop through the data (levels) 
                        foreach (var level in levels)
                        {
                            //Use String.split method to separate text file data
                            string levelName = level.Split(',')[0];
                            string levelValue = level.Split(',')[2];
                            
                            //Change level elevation values from string to double
                            double levMValue = Convert.ToDouble(levelValue);

                            //Change level elevation values unit from meters to feet
                            double levFValue = UnitUtils.Convert(levMValue, UnitTypeId.Meters, UnitTypeId.Feet);

                            //Create and name the levels
                            Level.Create(doc, levFValue).Name = levelName;
                            ElementId levelId = Level.Create(doc, levFValue).Id;


                            //Get viewFamilyTypeId

                            FilteredElementCollector viewTCollector = new FilteredElementCollector(doc);
                            FilteredElementCollector viewTypes = viewTCollector.OfCategory(BuiltInCategory.OST_Views).WhereElementIsElementType();
                          
                            //Get level Id

                            foreach (ElementType vType in viewTypes) 
                            {
                                if (vType.Name == "Floor Plan")
                                {

                                    ElementId viewFamilyTypeId =vType.Id;

                                    //Create view plans
                                    ViewPlan.Create(doc, viewFamilyTypeId,levelId);
                                }
                            }    



                          //Create RCPs
                        }



                        //Commit the levels Transaction
                        levTransaction.Commit();
                    }

                }

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C : \\";
                    openFileDialog.Filter = "csv files (*.csv) | *.csv | All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

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
                            //Create the collector and get the title block element ID
                            FilteredElementCollector collector = new FilteredElementCollector(doc);
                            ElementId titleBlockTypeId = collector.OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElementId();

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
                        sheetsTransaction.Dispose();

                    }

                }

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                message = e.Message;

                return Result.Failed;
            }

        }
    }
}
