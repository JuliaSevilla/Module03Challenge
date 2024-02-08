#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows;
using static Module03Challenge.Command1;

#endregion

namespace Module03Challenge
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

            ////0.OPTION 1 - HARD - CODED DATA PROVIDED BY MICHAEL
            ////1. get furniture data
            //List<string[]> furnitureTypeList = GetFurnitureTypes();
            //List<string[]> furnitureSetList = GetFurnitureSets();

            ////2. remove header row
            //furnitureTypeList.RemoveAt(0);
            //furnitureSetList.RemoveAt(0);

            //0.OPTION 2 - GET DATA FROM CSV FILES
            //0A. Get information from CSV file - Furniture
            // 0.1 Read csv file
            string furnitureData = "C:\\WIP\\RevitAddinBootcamp\\RAB_Module_03_Challenge_Files\\RAB_Module03_FurnitureCSV.csv";
            //another option of doing this     string furnitureData = @"C:\WIP\Revit Addin Bootcamp\RAB_Module_03_Challenge_Files";

            // 0.2 create list of string arrays from CSV 
            List<string[]> furnitureTypeList = new List<string[]>();

            // 0.3 Read text file datas. This will bring all lines into our array
            string[] furnitureArray = System.IO.File.ReadAllLines(furnitureData);

            // 0.4 loop through file data and put into list. We are creating a new array
            // (rowArray) but this time splitting each line
            foreach (string furnitureString in furnitureArray)
            {
                string[] rowArray = furnitureString.Split(',');
                furnitureTypeList.Add(rowArray);
            }
            // 0.5 Remove header row
            furnitureTypeList.RemoveAt(0);

            //0B. Get information from CSV file - Furniture
            // 0.1 Read csv file
            string furnitureSetData = "C:\\WIP\\RevitAddinBootcamp\\RAB_Module_03_Challenge_Files\\RAB_Module03_Furniture SetsCSV.csv";
            //another option of doing this     string furnitureData = @"C:\WIP\Revit Addin Bootcamp\RAB_Module_03_Challenge_Files";

            // 0.2 create list of string arrays from CSV 
            List<string[]> furnitureSetList = new List<string[]>();

            // 0.3 Read text file datas. This will bring all lines into our array
            string[] furnitureSetArray = System.IO.File.ReadAllLines(furnitureSetData);

            // 0.4 loop through file data and put into list. We are creating a new array
            // (rowArray) but this time splitting each line
            foreach (string furnitureSetString in furnitureSetArray)
            {
                string[] rowArray = furnitureSetString.Split(',');
                furnitureSetList.Add(rowArray);
            }
            // 0.5 Remove header row
            furnitureSetList.RemoveAt(0);

            //1. Create furniture data classes
            List<FurnitureType> furnitureTypes = new List<FurnitureType>();

            foreach (string[] curFurnTypeArray in furnitureTypeList)
            {
                FurnitureType curFurnType = new FurnitureType(curFurnTypeArray[0], curFurnTypeArray[1], curFurnTypeArray[2]);

                furnitureTypes.Add(curFurnType);

            }

            //2. Create furniture set classes
            List<FurnitureSet> furnitureSets = new List<FurnitureSet>();

            foreach (string[] curFurnSetArray in furnitureSetList)
            {
                FurnitureSet curFurnSet = new FurnitureSet(curFurnSetArray[0], curFurnSetArray[1], curFurnSetArray[2]);

                furnitureSets.Add(curFurnSet);
            }

            //3. Get Rooms from model - this are spatial elements (rooms, areas and spaces)
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            //4. Loop through rooms
            //variables within the using statement only live there
            using (Transaction t = new Transaction(doc))
            {
                int counter = 0;
                t.Start("Insert furniture into rooms");

                foreach (SpatialElement curRoom in collector)
                {

                    //Room location point
                    LocationPoint roomPoint = curRoom.Location as LocationPoint;
                    XYZ insPoint = roomPoint.Point as XYZ;

                    string furnSet = GetParameterValueAsString(curRoom, "Furniture Set");

                    // loop through furniture set data
                    foreach (FurnitureSet curSet in furnitureSets)
                    {
                        if(curSet.Set == furnSet)
                        {
                            foreach(string furnItem in curSet.Furniture)
                            {
                                foreach (FurnitureType curType in furnitureTypes)
                                {
                                    // make sure the strings we are comparing are exactly the same. trim get rid of spaces at the begininning or the end.
                                    if(furnItem.Trim() == curType.Name)
                                    {
                                        // Michael has it as "familyName" rather than RevitFamilyName
                                        FamilySymbol curFs = GetFamilySymbolByName(doc, curType.RevitFamilyName, curType.TypeName));
                                        // activate the family symbol if it is not activated
                                        // this needs to be within the transaction
                                        if(curFs != null)
                                        {
                                            if(curFs.IsActive == false)
                                            {
                                                curFs.Activate();
                                            }

                                            FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curFs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                            counter++;
                                        }
                                    }
                                }

                                SetParameterValue(curRoom, "Furniture Count", curSet.GetFurnitureCount());
                            }
                        }
                    }

                    t.Commit();

                }

                TaskDialog.Show("Complete", $"Inserted {counter} furniture instances");

                return Result.Succeeded;
            }

        }
        // void as this method doesnt return a value
        private void SetParameterValue(Element curElem, string paranName, int value)
        {
            //another way of doing it
            //Parameter curParam = curElem.LookupParameter(paranName);
            //if(curParam  != null )

            foreach(Parameter curParam in curElem.Parameters)
            {
                //we need to add the "definition" to get the name.
                if(curParam.Definition.Name == paranName)
                {
                    curParam.Set(value);
                }
            }
        }

        private void SetParameterValue(Element curElem, string paranName, string value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paranName)
                {
                    curParam.Set(value);
                }
            }
        }

        private void SetParameterValue(Element curElem, string paranName, double value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paranName)
                {
                    curParam.Set(value);
                }
            }
        }
        private void SetParameterValue(Element curElem, string paranName, ElementId value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paranName)
                {
                    curParam.Set(value);
                }
            }
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, string revitFamilyName, object typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach(FamilySymbol curFS in collector)
            {
                if (curFS.FamilyName == revitFamilyName && curFS.Name == typeName)
                {
                    return curFS;                
                }
            }
            return null;
        }

        public class FurnitureType
        {
            public string FurnitureName { get; set; }
            public string RevitFamilyName { get; set; }
            public string RevitFamilyType { get; set; }
            public List<FurnitureSet> FurnitureTypeList { get; set; }
        }

        public class FurnitureSet
        {
            public string FurnitureSetName { get; set; }
            public string RoomType { get; set; }
            public string RevitFamilyName { get; set; }
            public List<FurnitureSet> FurnitureSetList { get; set; }
        }

        private string GetParameterValueAsString(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    return curParam.AsString();
                }
            }
            return null;
        }
        public int GetFurnitureCount()
        {
            return Furniture.Length;
        }

        ////0.OPTION 1 - HARD - CODED DATA PROVIDED BY MICHAEL
        //private List<string[]> GetFurnitureTypes()
        //{
        //    List<string[]> returnList = new List<string[]>();
        //    returnList.Add(new string[] { "Furniture Name", "Revit Family Name", "Revit Family Type" });
        //    returnList.Add(new string[] { "desk", "Desk", "60in x 30in" });
        //    returnList.Add(new string[] { "task chair", "Chair-Task", "Chair-Task" });
        //    returnList.Add(new string[] { "side chair", "Chair-Breuer", "Chair-Breuer" });
        //    returnList.Add(new string[] { "bookcase", "Shelving", "96in x 12in x 84in" });
        //    returnList.Add(new string[] { "loveseat", "Sofa", "54in" });
        //    returnList.Add(new string[] { "teacher desk", "Table-Rectangular", "48in x 30in" });
        //    returnList.Add(new string[] { "student desk", "Desk", "60in x 30in Student" });
        //    returnList.Add(new string[] { "computer desk", "Table-Rectangular", "48in x 30in" });
        //    returnList.Add(new string[] { "lab desk", "Table-Rectangular", "72in x 30in" });
        //    returnList.Add(new string[] { "lounge chair", "Chair-Corbu", "Chair-Corbu" });
        //    returnList.Add(new string[] { "coffee table", "Table-Coffee", "30in x 60in x 18in" });
        //    returnList.Add(new string[] { "sofa", "Sofa-Corbu", "Sofa-Corbu" });
        //    returnList.Add(new string[] { "dining table", "Table-Dining", "30in x 84in x 22in" });
        //    returnList.Add(new string[] { "dining chair", "Chair-Breuer", "Chair-Breuer" });
        //    returnList.Add(new string[] { "stool", "Chair-Task", "Chair-Task" });

        //    return returnList;
        //}

        //private List<string[]> GetFurnitureSets()
        //{
        //    List<string[]> returnList = new List<string[]>();
        //    returnList.Add(new string[] { "Furniture Set", "Room Type", "Included Furniture" });
        //    returnList.Add(new string[] { "A", "Office", "desk, task chair, side chair, bookcase" });
        //    returnList.Add(new string[] { "A2", "Office", "desk, task chair, side chair, bookcase, loveseat" });
        //    returnList.Add(new string[] { "B", "Classroom - Large", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
        //    returnList.Add(new string[] { "B2", "Classroom - Medium", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
        //    returnList.Add(new string[] { "C", "Computer Lab", "computer desk, computer desk, computer desk, computer desk, computer desk, computer desk, task chair, task chair, task chair, task chair, task chair, task chair" });
        //    returnList.Add(new string[] { "D", "Lab", "teacher desk, task chair, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, stool, stool, stool, stool, stool, stool, stool" });
        //    returnList.Add(new string[] { "E", "Student Lounge", "lounge chair, lounge chair, lounge chair, sofa, coffee table, bookcase" });
        //    returnList.Add(new string[] { "F", "Teacher's Lounge", "lounge chair, lounge chair, sofa, coffee table, dining table, dining chair, dining chair, dining chair, dining chair, bookcase" });
        //    returnList.Add(new string[] { "G", "Waiting Room", "lounge chair, lounge chair, sofa, coffee table" });

        //    return returnList;



        //}

    }
}
