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


public class Rooms3D
{

    // 3D Rooms Création
    [Transaction(TransactionMode.Manual)]
    public class Create3Drooms : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view;
            view = doc.ActiveView;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Double roomNbre = 0;

            FilteredElementCollector m_Collector = new FilteredElementCollector(doc);
            m_Collector.OfCategory(BuiltInCategory.OST_Rooms);
            IList<Element> m_Rooms = m_Collector.ToElements();
            
          
            //  Iterate the list and gather a list of boundaries
            foreach (Room room in m_Rooms)
            {
                roomNbre += 1;
                //if (roomNbre == 10) { break; }

                //  Avoid unplaced rooms
                if (room.Area > 1)
                {

                    String TempPath = System.IO.Path.GetTempPath();
                    String PluginPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\";

                    //m_FamDoc = m_App.NewFamilyDocument(("C:\\Documents and Settings\\All Users\\Application Data\\" + "Autodesk\\RAC 2011\\Imperial Templates\\Specialty Equipment.rft"));
                    String _conceptual_mass_template_path = PluginPath + "EquipementSpecialiseMetrique.rft";
                    String _family_path = TempPath + "p5wf-" + room.UniqueId.ToString() + ".rfa";
                    String _family_name = "p5wf-" + room.UniqueId.ToString();
                    String departName = "";

                    try
                    {
                        if (File.Exists(_family_path)) { File.Delete(_family_path); }
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("Error Clean family file : {0}", ex.Message);
                    }

                    Dictionary<string, string> udroom = RvtVa3c.Util.GetElementProperties(room, true);
                    foreach (KeyValuePair<string, string> p in udroom)
                    {
                        if (p.Key.ToUpper() == "DEPARTEMENT (RM)")
                        {
                            departName = p.Value;
                            break;
                        }
                    }


                    Document massDoc = app.NewFamilyDocument(_conceptual_mass_template_path);

                    using (Transaction txMass = new Transaction(massDoc))
                    {
                        txMass.Start("Create Mass");

                        // Found BBOX

                        BoundingBoxXYZ bb = room.get_BoundingBox(null);
                        XYZ pt = new XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, bb.Min.Z);
                        RvtVa3c.Va3cExportContext.PointInt ptInt = new RvtVa3c.Va3cExportContext.PointInt(pt, true);
                        long ptIntYplus = ptInt.Y + 500;
                        udroom.Add("labelPosition", ptInt.X + "," + ptIntYplus + "," + ptInt.Z);


                        // Create Parameter

                        FamilyManager familyMgr = massDoc.FamilyManager;
                        udroom.Add("area", room.Area.ToString());
                        udroom.Add("room", room.Number);
                        udroom.Add("center", "");


                        foreach (KeyValuePair<string, string> p in udroom)
                        {

                            String pKeyOK = p.Key.Replace("?", "");
                            FamilyParameter param = familyMgr.get_Parameter(p.Key);
                            if (null == param)
                            {
                                param = massDoc.FamilyManager.AddParameter(
                                  pKeyOK, BuiltInParameterGroup.PG_DATA,
                                  ParameterType.Text, true);
                            }

                            try
                            {
                                if (massDoc.FamilyManager.Types.Size == 0)
                                {
                                    massDoc.FamilyManager.NewType("Type 1");
                                }
                                familyMgr.Set(param, p.Value); // set the value
                            }
                            catch (Exception ex)
                            {
                                Debug.Print("Error Set Param room : {0}", ex.Message);
                            }

                        }


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

                        String testCoord = "";

                        //XYZ centerObj = new XYZ(0, 0, 0);

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


                                    testCoord += (curve.Length + " : " + curve.GetEndPoint(0).X.ToString() + "," + curve.GetEndPoint(0).Y.ToString() + "," + curve.GetEndPoint(0).Z.ToString()) + "\r\n";


                                    m_CurveArray.Append(curve);


                                }



                            }
                        }


                        // MessageBox.Show(testCoord);




                        //  Simple insertion point
                        XYZ pt1 = new XYZ(0, 0, 0);
                        //  Our normal point that points the extrusion directly up
                        XYZ ptNormal = new XYZ(0, 0, 100);
                        //  The plane to extrude the mass from
                        Plane m_Plane = app.Create.NewPlane(ptNormal, pt1);
                        // SketchPlane m_SketchPlane = m_FamDoc.FamilyCreate.NewSketchPlane(m_Plane);
                        SketchPlane m_SketchPlane = SketchPlane.Create(massDoc, m_Plane); // 2014

                        //  Need to add the CurveArray to the final requirement to generate the form
                        CurveArrArray m_CurveArArray = new CurveArrArray();
                        m_CurveArArray.Append(m_CurveArray);

                        try
                        {
                            //  Extrude the form
                            Extrusion m_Extrusion = massDoc.FamilyCreate.NewExtrusion(true, m_CurveArArray, m_SketchPlane, 1);
                        }
                        catch (Exception ex)
                        {
                            Debug.Print("Error Extrusion room : {0}", ex.Message);
                        }


                        try
                        {
                            //  m_Extrusion.Subcategory = m_Subcategory;
                        }
                        catch //(Exception ex)
                        {
                        }

                        txMass.Commit();

                    }



                    SaveAsOptions opt = new SaveAsOptions();
                    opt.OverwriteExistingFile = true;
                    massDoc.SaveAs(_family_path, opt);

                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Create FaceWall");

                        if (!doc.LoadFamily(_family_path))
                            throw new Exception("DID NOT LOAD FAMILY");

                        Family family = new FilteredElementCollector(doc)
                          .OfClass(typeof(Family))
                          .Where<Element>(x => x.Name.Equals(_family_name))
                          .Cast<Family>()
                          .FirstOrDefault();

                        FamilySymbol fs = doc.GetElement(
                          family.GetFamilySymbolIds().First<ElementId>())
                            as FamilySymbol;

                        if (!fs.IsActive)
                        {
                            fs.Activate();
                        }



                        // Create a family instance

                        Level level = doc.ActiveView.GenLevel;

                        FamilyInstance fi = doc.Create.NewFamilyInstance(
                          XYZ.Zero, fs, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        //IList<Parameter> parameters = family.GetOrderedParameters();


                        doc.Regenerate(); // required to generate the geometry!

                        tx.Commit();
                    }

                    

                }
            }


            Debug.Print("Rooms total : {0}", roomNbre);
            return Result.Succeeded;

        }

    }


}
