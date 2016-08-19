#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;
using DialogResult = System.Windows.Forms.DialogResult;
using System.Collections;
using System.Text;
using Autodesk.Revit.DB.Architecture;
using System.Diagnostics;
#endregion // Namespaces



public class Rooms3DdirectShape
{

    // 3D Rooms Création
    [Transaction(TransactionMode.Manual)]
    public class Creat3Drooms : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view;
            view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            FilteredElementCollector m_Collector = new FilteredElementCollector(doc);
            m_Collector.OfCategory(BuiltInCategory.OST_Rooms);
            IList<Element> m_Rooms = m_Collector.ToElements();
            Double roomNbre = 0;

            // Create Shema Data
            SchemaBuilder schemaBuilder = new SchemaBuilder(new Guid("620080CB-DA99-40DC-9415-E53F280AA1F0"));

            // allow anyone to read the object
            schemaBuilder.SetReadAccessLevel(
              AccessLevel.Public);

            // restrict writing to this vendor only
            schemaBuilder.SetWriteAccessLevel(
              AccessLevel.Vendor);

            // required because of restricted write-access
            schemaBuilder.SetVendorId("XXXX"); // same than the XML start file
            
            // create a field to store string
            Dictionary<string, string> listFiled = new Dictionary<string, string>();
            listFiled.Add("area", "");
            listFiled.Add("building", "");
            listFiled.Add("categorie", "");
            listFiled.Add("label", "");
            listFiled.Add("labelPosition", "");
            listFiled.Add("level", "");
            listFiled.Add("name", "");
            listFiled.Add("revit_id", "");
            listFiled.Add("room", "");
            listFiled.Add("site", "");
            listFiled.Add("styleCategory", "");
            listFiled.Add("team", "");
            listFiled.Add("type", "");
            listFiled.Add("workset", "");

            foreach (KeyValuePair<string, string> p in listFiled)
            {
                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(p.Key, typeof(String));
                fieldBuilder.SetDocumentation("Rooms infos");
            }


            schemaBuilder.SetSchemaName("Room3D");
            Schema schema = schemaBuilder.Finish(); // register the Schema object
            
            //  Iterate the list and gather a list of boundaries
            foreach (Room room in m_Rooms)
            {
                roomNbre += 1;
                
                //  Avoid unplaced rooms
                if (room.Area > 1)
                {

                    using (Transaction tr = new Transaction(doc))
                    {
                        tr.Start("Create Mass");

                        // Found BBOX

                        BoundingBoxXYZ bb = room.get_BoundingBox(null);
                        XYZ pt = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, bb.Min.Z);
                        RvtVa3c.Va3cExportContext.PointInt ptInt = new RvtVa3c.Va3cExportContext.PointInt(pt, true);
                        long ptIntYplus = ptInt.Y;


                        //  Get the room boundary
                        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

                        // a room may have a null boundary property:

                        int n = 0;

                        if (null != boundaries)
                        {
                            n = boundaries.Count; // 2012
                        }


                        //  The array of boundary curves
                        CurveArray m_CurveArray = new CurveArray();
                        //  Iterate to gather the curve objects
                        List<Curve> profile = new List<Curve>();
                      
                        if (0 < n)
                        {
                            int iBoundary = 0, iSegment;

                            foreach (IList<BoundarySegment> b in boundaries) // 2012
                            {
                                ++iBoundary;
                                iSegment = 0;
                                foreach (BoundarySegment s in b)
                                {
                                    ++iSegment;
                                    Curve curve = s.GetCurve(); // 2016

                                    profile.Add(curve); //add shape for instant object

                                }

                            }
                        }


                        try
                        {
                            
                            // Add Direct Shape
                            CurveLoop curveLoop = CurveLoop.Create(profile);
                            List<CurveLoop> curveLoopList = new List<CurveLoop>();
                            curveLoopList.Add(curveLoop);

                            SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                            
                            //  Simple insertion point
                            XYZ pt1 = new XYZ(0, 0, 0);
                            //  Our normal point that points the extrusion directly up
                            XYZ ptNormal = new XYZ(0, 0, 100);
                           
                            // Solid roomSolid = GeometryCreationUtilities.CreateRevolvedGeometry(frame, new CurveLoop[] { curveLoop }, 0, 2 * Math.PI, options);
                            Solid roomSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, ptNormal, 1);
                            
                            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel),
                                                                            "room3Dds",
                                                                             room.UniqueId.ToString());

                            ds.SetShape(new GeometryObject[] { roomSolid });
                            


                            // Add data
                            
                            Dictionary<string, string> udroom = RvtVa3c.Util.GetElementProperties(room, true);
                            Dictionary<string, string> udroomNew = new Dictionary<string, string>();
                            
                            foreach (KeyValuePair<string, string> p in udroom)
                            {
                                string dKey = p.Key;
                                string dValue = p.Value;

                                switch (dKey.ToUpper())
                                {
                                    case "ROOM":
                                        dKey = "room";
                                        break;
                                    case "NIVEAU":
                                        dKey = "level";
                                        break;
                                    case "NOM":
                                        dKey = "name";
                                        if (dValue == "") { dValue = "D00_000"; }
                                        break;
                                    case "SOUS-PROJET":
                                        dKey = "workset";
                                        break;
                                    default:
                                        dKey = "";
                                        break;
                                }

                                if (dKey != "")
                                {
                                    udroomNew.Add(dKey, dValue);
                                }

                            }

                            udroomNew.Add("area", room.Area.ToString());
                            udroomNew.Add("room", room.Number);
                            udroomNew.Add("revit_id", room.UniqueId.ToString());


                            // Create Parameter

                            try
                            {
                                // create an entity (object) for this schema (class)
                                Entity entity = new Entity(schema);

                                foreach (KeyValuePair<string, string> p in udroomNew)
                                {

                                    // get the field from the schema
                                    Field fieldSpliceLocation = schema.GetField(p.Key);
                                    entity.Set<String>(fieldSpliceLocation, p.Value.ToString());  // , DisplayUnitType.DUT_METERS); // set the value for this entity
                                    ds.SetEntity(entity); // store the entity in the element

                                }
                            }
                            catch (Exception s)
                            {
                                Console.WriteLine(s.Message);
                            }


                            tr.Commit();


                            // get the data back from the wall
                            Entity retrievedEntity = ds.GetEntity(schema);

                            String retrievedData = retrievedEntity.Get<String>(
                              schema.GetField("area")); //,DisplayUnitType.DUT_METERS);


                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);

                        }


                    }



                }
            }


            Debug.Print("Rooms total : {0}", roomNbre);
            return Result.Succeeded;

        }

    }


}
