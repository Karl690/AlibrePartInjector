using AlibreX;
using System;
using System.Runtime.InteropServices;

namespace AlibreConsoleApp
{
    public class Program
    {
        //ref string ParttoAddToAssemblyFullPath = "";
        private static IAutomationHook hook;                  //Holds Alibre Automation hook object
        private static IADRoot rootObj;                      //Holds Alibre Root object
        private static IADOccurrence objADOccurrence;        //Holds Occurrence object
        private static IADDesignSession DesignSession;       //Holds Alibre Design Session object
        private static IADPartSession objPartSession;        //Holds Alibre Part Session object
        private static bool insertEnabled = true;

        public static void Main(string[] args)
        {
            frmPartInsert_Load();
            addPartToAssembly(args[0]);
        }

        //This function gets the automation hook for the running instance of Alibre.
        //If there is any Assembly Session open, then the 'Insert part and Save Assembly' button gets enabled.
        //If there is no Assembly Session open, this button remains disabled.
        private static void frmPartInsert_Load()
        {
            try
            {
                hook = (IAutomationHook)Marshal.GetActiveObject("AlibreX.AutomationHook"); //Gets the automation hook for the running instance of Alibre
                rootObj = (IADRoot)hook.Root;
                insertEnabled = true;
            }
            catch
            {
                Console.WriteLine("Launch Alibre Design and restart this application");
                insertEnabled = false;
            }
        }

        //This function inserts a part into the first Assembly Session and then Saves it to C:\ Drive
        private static void btnInsert_Click(string filePath)
        {
            addPartToAssembly(filePath/*@"Y:\Master Components\103356\103356 high pressure printer rotary print arm.AD_PRT"*/);
        }

        public static void addPartToAssembly(string fileNameToLoad)
        {
            IADAssemblySession objAssmSession;     //Holds Alibre Assembly Session Object
            IADOccurrence objADRootOccurrence;     //Holds Root Occurrence of the Assembly
            IADOccurrences objADOccurrences;       //Holds all Occurrences of the Assembly
            string destinationString;              //Holds the location where the File gets Saved
            bool flag = true;


            if (rootObj == null || rootObj.Sessions == null)  //Exit if for some reason an instance of Alibre Design could not be found
                return;

            if (rootObj.Sessions.Count > 0)         // If there is atleast one workspace open
            {
                foreach (IADSession objSession in rootObj.Sessions)
                {
                    if ((objSession.SessionType == ADObjectSubType.AD_ASSEMBLY) && flag)     //If there is atleast one Assembly open
                    {
                        objAssmSession = (IADAssemblySession)objSession;                                 //part is inserted into that assembly
                        flag = false;

                        Console.WriteLine("Inserting Part into the Assembly...");

                        //Get Root Occurrence from Assembly Session
                        objADRootOccurrence = objAssmSession.RootOccurrence;

                        //Get Occurrences collection from Root Occurance
                        objADOccurrences = objADRootOccurrence.Occurrences;

                        //Holds Geometry Factory and et Geometry Factory from Session object
                        IADGeometryFactory objADGeometryFactory = objSession.GeometryFactory;

                        //Populate the Transformation Array with the following Data for Back View
                        double[] adblTransformationArrayData = new double[16] {1,0,0,0,
                                                                               0,1,0,0,
                                                                               0,0,1,0,
                                                                               0,0,0,1};

                        //Create Transformation
                        Array sysArray = adblTransformationArrayData;
                        IADTransformation objADTransformation = objADGeometryFactory.CreateTransform(ref sysArray);

                        //Add an Empty Part as Occurrence
                        //objADOccurrence = objADOccurrences.AddEmptyPart("BlockWithHole", false, objADTransformation);
                        object ParttoAddToAssemblyFullPath = fileNameToLoad;
                        objADOccurrence = objADOccurrences.Add(ref ParttoAddToAssemblyFullPath, objADTransformation);
                        //Set Design Session to be the empty Part's Design Session that was just added to the assembly
                        DesignSession = objADOccurrence.DesignSession;

                        //Set Part Session to be the empty Part's Design Session that was just added to the assembly
                        objPartSession = (IADPartSession)DesignSession;

                        //Call to CreateFeatures method to add features to the empty part inserted into the assembly
                        //CreateFeatures();

                        //lblStatus.Text = "Part inserted successfully into " + objSession.Name;

                        ////Saves the Assembly with the Part to the location specified
                        //lblStatus.Text = "Saving assembly on C:\\ Drive...";
                        //destinationString = "C://";
                        //object saveLocation = (object)destinationString;
                        //try
                        //{
                        //    objSession.SaveAs(ref saveLocation, objSession.Name);
                        //    lblStatus.Text = "Assembly is saved successfully on C:\\";
                        //    btnInsert.Enabled = false;
                        //}
                        //catch
                        //{
                        //    lblStatus.Text = "Assembly in use, unable to save.";
                        //}

                        break;
                    }
                    else
                        Console.WriteLine("Please open any Assembly");

                }
            }
            else        //If there is no assembly open
                Console.WriteLine("Please open any Assembly");
        }

        private static void CreateFeatures()
        {
            IADDesignPlanes allPlanes;             //Holds Design Planes
            IADDesignPlane refPlane;               //Holds Design Plane
            IADSketch objPlaneSketch;              //Holds the Reference Sketch
            IADSketchFigures objADSketchFigures;   //Holds all Sketch Figures
            IADPartFeatures objFeatures;           //Holds all Part Features
            IADPartFeature objExtrudeBossFeature;  //Holds the Extrusion Feature

            allPlanes = DesignSession.DesignPlanes;  //Get all Planes in the Part
            refPlane = allPlanes.Item("XY-Plane");   //Get XY Plane
            objPlaneSketch = objPartSession.Sketches.AddSketch(null, refPlane, "Sketch1"); //Add Sketch to XY Plane
            objADSketchFigures = objPlaneSketch.Figures; //Get the Sketch added to XY Plane

            //The following calls sketch a Rectangle and a Circle in the XY Plane
            objPlaneSketch.BeginChange();
            objPlaneSketch.Figures.AddRectangle(-10, -10, 10, 10);
            objPlaneSketch.Figures.AddCircle(0, 0, 5);
            objPlaneSketch.EndChange();

            objPartSession.Sketches.Item(0).Name = "Sketch1";
            objPlaneSketch = objPartSession.Sketches.Item(0); //Name the Sketch as Sketch1
            objFeatures = objPartSession.Features;
            //Adds the Extrusion feature using the Sketch created above
            objExtrudeBossFeature = (IADPartFeature)objFeatures.AddExtrudedBoss(objPlaneSketch, (object)5,
                                ADPartFeatureEndCondition.AD_MID_PLANE, null, null, 0, ADDirectionType.AD_ALONG_NORMAL, null,
                                null, false, (object)0, false, "BlockWithHoleFeature", "ExtrudeDepth", "ExtrudeAngle");
        }
    }
}
