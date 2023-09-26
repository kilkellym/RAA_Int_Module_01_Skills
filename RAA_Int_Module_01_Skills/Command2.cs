#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAA_Int_Module_01_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 01. Create schedule
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create schedule");

                ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
                newSchedule.Name = "My Room Schedule";

                // 02a. Get parameters for fields
                FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
                doorCollector.OfCategory(BuiltInCategory.OST_Rooms);
                doorCollector.WhereElementIsNotElementType();

                Element doorInst = doorCollector.FirstElement();

                Parameter roomNumParam = doorInst.LookupParameter("Number");
                Parameter roomNameParam = doorInst.LookupParameter("Name");
                Parameter roomLevelParam = doorInst.LookupParameter("Level");
                Parameter roomAreaParam = doorInst.get_Parameter(BuiltInParameter.ROOM_AREA);

                Parameter level1 = doorInst.get_Parameter(BuiltInParameter.LEVEL_NAME);
                Parameter level2 = doorInst.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID);


                // 02b. Create fields
                ScheduleField roomNumField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNumParam.Id);
                ScheduleField roomLevelField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomLevelParam.Id);
                ScheduleField roomNameField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
                ScheduleField roomAreaField = newSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);
                //ScheduleField roomLevel2Field = newSchedule.Definition.AddField(ScheduleFieldType.Instance, level2.Id);

                // filter schedule by department
                Element myLevel = GetLevelByName(doc, "02 - Floor");
                
                //ScheduleFilter deptFilter = new ScheduleFilter(roomAreaField.FieldId, ScheduleFilterType.GreaterThan, 1000);
                //newSchedule.Definition.AddFilter(deptFilter);

                var RoomSchedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(curViewScheduleInDoc => curViewScheduleInDoc.Name == "Room Finish Schedule");

                var roomInstance = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)

                                                        .WhereElementIsNotElementType()

                                                             .FirstOrDefault();

                // Get Area fiel
                List<ScheduleFieldId> fieldList = RoomSchedule.Definition.GetFieldOrder().ToList();
                Parameter areaParam = roomInstance.get_Parameter(BuiltInParameter.ROOM_AREA);

                foreach (ScheduleFieldId curId in fieldList)
                {
                    ScheduleField curField = RoomSchedule.Definition.GetField(curId);

                    if(curField.ParameterId == roomAreaParam.Id)
                    {
                        curField.DisplayType = ScheduleFieldDisplayType.Totals;
                    }
                }
                
               
                t.Commit();
            }


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
        public static Level GetLevelByName(Document curDoc, string levelName)
        {
            FilteredElementCollector levelCollector = new FilteredElementCollector(curDoc);
            levelCollector.OfCategory(BuiltInCategory.OST_Levels);
            levelCollector.OfClass(typeof(Level)).ToElements();

            //loop through levels and find match for levelName argument
            foreach (Level tmpLevel in levelCollector)
            {
                if (tmpLevel.Name == levelName)
                {
                    return tmpLevel;
                }
            }

            return null;
        }
    }
}
