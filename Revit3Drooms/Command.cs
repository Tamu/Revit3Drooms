#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using System.Diagnostics;
using Autodesk.Revit.DB.ExtensibleStorage;
#endregion // Namespaces

namespace Revit3Drooms
{

    // 3D Rooms Création
    [Transaction(TransactionMode.Manual)]
    public class Create3Drooms : IExternalCommand
    {

        public static Guid _schemaGuid = new Guid("420080CB-DA99-22DC-9415-E53F280AA1F0");


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view;
            view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;


            // Deleting existing DirectShape
            // get ready to filter across just the elements visible in a view 
            FilteredElementCollector coll = new FilteredElementCollector(doc, view.Id);
            coll.OfClass(typeof(DirectShape));
            IEnumerable<DirectShape> DSdelete = coll.Cast<DirectShape>();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Delete elements");

                try
                {

                    foreach (DirectShape ds in DSdelete)
                    {
                        ICollection<ElementId> ids = doc.Delete(ds.Id);
                    }

                    tx.Commit();
                }
                catch (ArgumentException)
                {
                    tx.RollBack();
                }
            }


            // Delete Shema
            try
            {
                using (Transaction trans = new Transaction(doc))
                {

                    IList<Schema> schemasIn = Schema.ListSchemas();
                    foreach (Schema s in schemasIn)
                    {
                        Console.WriteLine("In > " + s.SchemaName + " : " + s.GUID);
                    }

                    trans.Start("Delete Schema");
                    Schema schemaOld = Schema.Lookup(_schemaGuid);
                    Schema.EraseSchemaAndAllEntities(schemaOld, false);
                    trans.Commit();

                    IList<Schema> schemasOut = Schema.ListSchemas();
                    foreach (Schema s in schemasOut)
                    {
                        Console.WriteLine("Out > " + s.SchemaName + " : " + s.GUID);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            // Def Site + Building
            double roomNbre = 0;
            string filename = doc.Title;


            FilteredElementCollector m_Collector = new FilteredElementCollector(doc);
            m_Collector.OfCategory(BuiltInCategory.OST_Rooms);
            IList<Element> m_Rooms = m_Collector.ToElements();

            // Create Shema Data
            SchemaBuilder schemaBuilder = new SchemaBuilder(_schemaGuid);

            // allow anyone to read the object
            schemaBuilder.SetReadAccessLevel(
              AccessLevel.Public);

            // restrict writing to this vendor only
            schemaBuilder.SetWriteAccessLevel(
              AccessLevel.Vendor);

            // required because of restricted write-access
            schemaBuilder.SetVendorId("P5RD");

            // create a field to store string
            Dictionary<string, string> listFiled = new Dictionary<string, string>();
            listFiled.Add("accessories", "");
            listFiled.Add("area", "");
            listFiled.Add("booking", "");
            listFiled.Add("building", "");
            listFiled.Add("categorie", "");
            listFiled.Add("center", "");
            listFiled.Add("department", "");
            listFiled.Add("division", "");
            listFiled.Add("label", "");
            listFiled.Add("labelPosition", "");
            listFiled.Add("level", "");
            listFiled.Add("name", "");
            listFiled.Add("occupationType", "");
            listFiled.Add("revit_id", "");
            listFiled.Add("room", "");
            listFiled.Add("sap", "");
            listFiled.Add("service", "");
            listFiled.Add("site", "");
            listFiled.Add("styleCategory", "");
            listFiled.Add("team", "");
            listFiled.Add("type", "");
            listFiled.Add("videoConferencing", "");
            listFiled.Add("workset", "");
            listFiled.Add("update", "");

            foreach (KeyValuePair<string, string> p in listFiled)
            {
                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(p.Key, typeof(String));
                fieldBuilder.SetDocumentation("Rooms infos");

            }


            schemaBuilder.SetSchemaName("testRoom");
            Schema schema = schemaBuilder.Finish(); // register the Schema object

            //  Iterate the list and gather a list of boundaries
            foreach (Room room in m_Rooms)
            {

                //  Avoid unplaced rooms
                if (room.Area > 1)
                {


                    String _family_name = "testRoom-" + room.UniqueId.ToString();

                    using (Transaction tr = new Transaction(doc))
                    {
                        tr.Start("Create Mass");

                        // Found BBOX

                        BoundingBoxXYZ bb = room.get_BoundingBox(null);
                        XYZ pt = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, bb.Min.Z);


                        //  Get the room boundary
                        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions()); // 2012


                        // a room may have a null boundary property:
                        int n = 0;

                        if (null != boundaries)
                        {
                            n = boundaries.Count;
                        }


                        //  The array of boundary curves
                        CurveArray m_CurveArray = new CurveArray();
                        //  Iterate to gather the curve objects

                       
                        TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
                        builder.OpenConnectedFaceSet(true);

                        // Add Direct Shape
                        List<CurveLoop> curveLoopList = new List<CurveLoop>();

                        if (0 < n)
                        {
                            int iBoundary = 0, iSegment;

                            foreach (IList<BoundarySegment> b in boundaries) // 2012
                            {
                                List<Curve> profile = new List<Curve>();
                                ++iBoundary;
                                iSegment = 0;

                                foreach (BoundarySegment s in b)
                                {
                                    ++iSegment;
                                    Curve curve = s.GetCurve(); // 2016

                                    profile.Add(curve); //add shape for instant object

                                }


                                try
                                {
                                    CurveLoop curveLoop = CurveLoop.Create(profile);
                                    curveLoopList.Add(curveLoop);
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine(ex.Message);
                                }

                            }



                        }



                        try
                        {

                            SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);

                            Frame frame = new Frame(pt, XYZ.BasisX, -XYZ.BasisZ, XYZ.BasisY);

                            //  Simple insertion point
                            XYZ pt1 = new XYZ(0, 0, 0);
                            //  Our normal point that points the extrusion directly up
                            XYZ ptNormal = new XYZ(0, 0, 100);
                            //  The plane to extrude the mass from
                            Plane m_Plane = app.Create.NewPlane(ptNormal, pt1);
                            // SketchPlane m_SketchPlane = m_FamDoc.FamilyCreate.NewSketchPlane(m_Plane);
                            SketchPlane m_SketchPlane = SketchPlane.Create(doc, m_Plane); // 2014

                            Solid roomSolid;

                            // Solid roomSolid = GeometryCreationUtilities.CreateRevolvedGeometry(frame, new CurveLoop[] { curveLoop }, 0, 2 * Math.PI, options);
                            roomSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoopList, ptNormal, 1);

                            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel), "3drooms", _family_name);

                            ds.SetShape(new GeometryObject[] { roomSolid });


                            // Add data
                            Dictionary<string, string> udroom = Revit3Drooms.Util.GetElementProperties(room, true);
                            Dictionary<string, string> udroomNew = new Dictionary<string, string>();

                            foreach (KeyValuePair<string, string> p in udroom)
                            {
                                if (p.Key.ToLower() == "name")
                                {
                                    udroomNew.Add(p.Key, p.Value);
                                }
                                else if (p.Key.ToLower() == "niveau")
                                {
                                    udroomNew.Add("level", p.Value);
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
                                    entity.Set<String>(fieldSpliceLocation, p.Value.ToString());  // set the value for this entity
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
                              schema.GetField("building"));

                            roomNbre += 1;

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);

                        }

                    }

                }
            }

            Debug.Print("Rooms total : {0}", roomNbre + "/" + m_Rooms.Count.ToString());
            Console.WriteLine("Total Rooms : " + roomNbre + "/" + m_Rooms.Count.ToString());
            return Result.Succeeded;

        }

    }
}