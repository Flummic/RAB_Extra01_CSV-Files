#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace Extra01_CSV_Files
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            // 1. Declare variables
            string levelPath = "D:\\___ NBM ___\\7. Self-Actualisation\\Revit-Addin-Academy\\Extra01_CSV-Files\\RAB_Bonus_Levels.csv"; // Hardcoded File-Path (with double-slashes because of C#)

            // 2. Create a list of string arrays for CSV data (Data Container)
            List<string[]> levelData = new List<string[]>(); // created the list "levelData" which will contain string-arrays

            // 3. read text file datas
            string[] levelArray = System.IO.File.ReadAllLines(levelPath);

            // 4. loop though file data and put into list
            foreach (string levelString in levelArray)
            {
                string[] rowArray = levelString.Split(','); // create a new Array ("rowArray") and fill it with the current line ("levelString") split by a comma 
                levelData.Add(rowArray); // add the new array (which includes the separated parts of the CSV-line (In levelArray) to the list ("levelData")
            }

            // 5. remove Header-Row from the list
            levelData.RemoveAt(0); // removes the first "row" (item) from the list

            // 6. create transaction (because we are about to change the Revit-Model)
            Transaction t = new Transaction(doc);
            t.Start("Create Levels from CSV");

            // 7. loop through level data
            int counter = 0;
            foreach (string[] currentLevelData in levelData)
            {
                // 8. create height variables
                double heightFeet = 0;
                double heightMeters = 0;

                // 9. get height and convert from string to double (The Conversion is neccessary becaus we read the csv-lines as strings)
                bool convertFeet = double.TryParse(currentLevelData[1], out heightFeet); // try to convert the second array-item of the currentLevelData-String-Array into a double and if it works (bool = True) overwrtie heightFeet with the new value
                bool convertMeters = double.TryParse(currentLevelData[2], out heightMeters);

                // 10. if using metric, convert meters to feet (because the Revit-API works with feet)
                double heightMetersConvert = heightMeters * 3.28084;
                double heightMetersConvert2 = UnitUtils.ConvertToInternalUnits(heightMeters, UnitTypeId.Meters); // alternative Method of doing the conversion

                // 11. create Level and rename it
                Level currentLevel = Level.Create(doc, heightFeet); // create a Level at the current height in Feet
                currentLevel.Name = currentLevelData[0]; // rename the level to the Name taken from the CSV (index 0 in the CurrentLevelData-String-Array

                // 12. increment counter (to count how many Levels were created by the addin
                counter++;
            }

            // 13. Commit and Dispose the transaction
            t.Commit();
            t.Dispose();

            // 14. Prompt the user with a dialog that shows how many Levels were created from which file
            TaskDialog.Show("Complete", "Created " + counter.ToString() + " levels from the CSV-File located at: " + levelPath);

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
