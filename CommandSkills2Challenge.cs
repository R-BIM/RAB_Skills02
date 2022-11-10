#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Security.Cryptography.X509Certificates;


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
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            try
            {
                
              //Create the active document levels form a CSV file

                // Get the CSV file for the levels
                string LevfilePath = @"C:\Users\rafik\Downloads\RAB_Session_02_Challenge_Levels.csv";

                if (File.Exists(LevfilePath))
                {

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

                    }
                    //Commit the levels Transaction
                    levTransaction.Commit();



                    //Create the active document sheets
                    // get the CSV file for the sheets
                    string ShtsfilePath = @"C:\Users\rafik\Downloads\RAB_Session_02_Challenge_Sheets.csv";
                    if (File.Exists(ShtsfilePath))
                    {

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
                            //Create the collector
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

                    }

                }

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                message= e.Message; 

                return Result.Failed;
            }           

        }



        const double _inchToMm = 25.4;
        const double _footToMm = 12 * _inchToMm;
        const double mToInch = 39.37;
        const double mToFoot = 3.28;

        /// <summary>
        /// Convert a given length in feet to millimetres.
        /// </summary>
        public static double FootToM(int length)
        {
            return length * mToFoot;
        }


    }
}
