using System;
using System.Windows;
using INFITF;
using MECMOD;
using PARTITF;
using HybridShapeTypeLib;
using CATMat;
using ProductStructureTypeLib;

namespace Schraubenprogramm

{
    internal class CatiaControl
    {
        INFITF.Application stg_catiaApp;
        MECMOD.PartDocument stg_catiaPart;
        MECMOD.Sketch stg_catiaSchaftProfil;
        MECMOD.Sketch stg_catiaKopfProfil;
        MECMOD.Sketch stg_catiaTasche;
        MECMOD.Sketch stg_NG_CatiaSchaftProfil;
        MECMOD.Sketch stg_NG_catiaKopfProfil;
        
        Factory2D F2D;
        ShapeFactory SF2D;
        HybridShapeFactory HSF2D;
        Pad catSchaftBlock;
        Body catBody;
        Part catPart;
        Sketches catSketches;
        OriginElements catOriginElements;
        HybridShapePlaneOffset hybridShapePlaneOffset1;


        //Product einbindungen
        ProductDocument productDocuments1;
        PartDocument partDocument;
        Part part1;
        Part part2;
        Part part3;
        MECMOD.Sketch stg_cat_Profil1;
        MECMOD.Sketch stg_cat_Profil2;
        MECMOD.Sketch stg_cat_ProfilMutterBohrung;
        MECMOD.Sketch stg_cat_KopfProfilx;
        MECMOD.Sketch stg_cat_UnterlegProfil;
        MECMOD.Sketch stg_cat_UnterlegBohrung;
        Factory2D catFactory2D;
        Factory2D catFac;


        //Achsensysteme erzeugen
        #region Achsensysteme erzeugen 
        //Achsensystem für Anpassung Schaft
        private void ErzeugeAchsensystem()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_catiaSchaftProfil.SetAbsoluteAxisData(array);
        }
        //Achsensystem für normgerechter Schaft
        private void NG_ErzeugeAchsensystem()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_NG_CatiaSchaftProfil.SetAbsoluteAxisData(array);
        }
        //Achsensystem für normgerechter Kopf
        private void NG_ErzeugeAchsensystem2()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_NG_catiaKopfProfil.SetAbsoluteAxisData(array);
        }
        private void ErzeugeAchsensystemX()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_cat_Profil1.SetAbsoluteAxisData(array);
        }
        private void ErzeugeAchsensystemY()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_cat_Profil2.SetAbsoluteAxisData(array);
        }
        private void ErzeugeAchsensystemS()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_cat_KopfProfilx.SetAbsoluteAxisData(array);
        }
        private void ErzeugeAchsensystemZ()
        {
            object[] array = new object[]
            {
                0.0, 0.0, 0.0,
                0.0, 1.0, 0.0,
                0.0, 0.0, 1.0
            };
            stg_cat_UnterlegBohrung.SetAbsoluteAxisData(array);
        }
        #endregion

        //Catia Läuft
        #region Catia läuft
        internal bool CatiaLäuft()
        {


            try
            {
                object CatiaObject = System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
                stg_catiaApp = (INFITF.Application)CatiaObject;
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }
        #endregion


        //Part erzeugen und öffnen
        #region Part erzeugen
        internal void ErzeugePart()
        {
            Documents CatDocument = (Documents)stg_catiaApp.Documents;
            stg_catiaPart = (PartDocument)CatDocument.Add("Part");
        }
        #endregion

        //ANPASSUNG
        //SCHAFT
        //Skizze für Schaft erzeugen
        #region Skizze erzeugen
        internal void ErzeugeSchaftSkizze()        //Schaft Skizze erzeugen 
        {
            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;
            HybridBodies HybridBodySchaft1 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodySchaft1;


            try
            {
                catHybridBodySchaft1 = HybridBodySchaft1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }


            catSketches = catHybridBodySchaft1.HybridSketches;
            catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference catref = (Reference)catOriginElements.PlaneYZ;
            stg_catiaSchaftProfil = catSketches.Add(catref);

            ErzeugeAchsensystem();

            stg_catiaPart.Part.Update();


        }
        #endregion

        //Schraubenschaft mit Gewinde erzeugen
        #region Schraubenschaft definieren
        //Skizze und Block für Schaft definieren
        #region Skizze und Block
        internal void ErzeugeSchaftBlock(Schraube Schraubeneigenschaft)
        {
            catPart = stg_catiaPart.Part;
            Bodies bodies = catPart.Bodies;
            catBody = catPart.MainBody;

            catPart.InWorkObject = catPart.MainBody;

            stg_catiaSchaftProfil.set_Name("Kreis");

            F2D = stg_catiaSchaftProfil.OpenEdition();

            Point2D Mittelpunkt = F2D.CreatePoint(0, 0);
            Circle2D Kreis = F2D.CreateCircle(0, 0, Schraubeneigenschaft.innenradius, 0, 0);
            Kreis.CenterPoint = Mittelpunkt;

            stg_catiaSchaftProfil.CloseEdition();

            catSchaftBlock = SF2D.AddNewPad(stg_catiaSchaftProfil, Schraubeneigenschaft.laenge);

            

           

            catPart.Update();

        }
        #endregion 
        //Gewinde erzeugen
        #region Gewinde
        //Helix erzeugen
        internal void ErzeugeGewindeHelix(Schraube Schraubeneigenschaft, string gewindeart)
        {
            Double P = Schraubeneigenschaft.steigung;
            Double Ri = Schraubeneigenschaft.innenradius;
            HSF2D = (HybridShapeFactory)catPart.HybridShapeFactory;

            Sketch myGewinde = null;

            if (gewindeart == "Withworth")
            {
                myGewinde = makeGewindeSkizzeWithworth(Schraubeneigenschaft);
            }
            else if (gewindeart == "Trapez")
            {
                myGewinde = makeGewindeSkizzeTrapez(Schraubeneigenschaft);
            }
            else if (gewindeart == "Sägen")
            {
                myGewinde = makeGewindeSkizzeSägen(Schraubeneigenschaft);
            }
            else if (gewindeart == "MetrsichesGewinde")
            {
                myGewinde = makeGewindeSkizzeWithworth(Schraubeneigenschaft);
            }

            HybridShapeDirection HelixDir = HSF2D.AddNewDirectionByCoord(1, 0, 0);
            Reference RefHelixDir = catPart.CreateReferenceFromObject(HelixDir);

            HybridShapePointCoord HelixStartpunkt = HSF2D.AddNewPointCoord(0, Ri, 0);
            Reference RefHelixStartpunkt = catPart.CreateReferenceFromObject(HelixStartpunkt);

            HybridShapeHelix Helix = HSF2D.AddNewHelix(RefHelixDir, false, RefHelixStartpunkt, P, Schraubeneigenschaft.gewindeLaenge, false, 0, 0, false);

            Reference RefHelix = catPart.CreateReferenceFromObject(Helix);
            Reference RefmyGewinde = catPart.CreateReferenceFromObject(myGewinde);

            catPart.Update();

            catPart.InWorkObject = catBody;

            OriginElements catOriginElements = this.stg_catiaPart.Part.OriginElements;
            Reference RefmyPlaneZX = (Reference)catOriginElements.PlaneZX;

            Sketch ChamferSkizze = catSketches.Add(RefmyPlaneZX);
            catPart.InWorkObject = ChamferSkizze;
            ChamferSkizze.set_Name("Fase");

            double H_links = Ri;
            double V_links = 0.65 * P;

            double H_rechts = Ri;
            double V_rechts = 0;

            double H_unten = Ri - 0.65 * P;
            double V_unten = 0;

            Factory2D catFactory2D3 = ChamferSkizze.OpenEdition();

            Point2D links = catFactory2D3.CreatePoint(H_links, V_links);
            Point2D rechts = catFactory2D3.CreatePoint(H_rechts, V_rechts);
            Point2D unten = catFactory2D3.CreatePoint(H_unten, V_unten);

            Line2D Oben = catFactory2D3.CreateLine(H_links, V_links, H_rechts, V_rechts);
            Oben.StartPoint = links;
            Oben.EndPoint = rechts;

            Line2D hypo = catFactory2D3.CreateLine(H_links, V_links, H_unten, V_unten);
            hypo.StartPoint = links;
            hypo.EndPoint = unten;

            Line2D seite = catFactory2D3.CreateLine(H_unten, V_unten, H_rechts, V_rechts);
            seite.StartPoint = unten;
            seite.EndPoint = rechts;

            ChamferSkizze.CloseEdition();

            catPart.InWorkObject = catBody;
            catPart.Update();

            Groove myChamfer = SF2D.AddNewGroove(ChamferSkizze);
            myChamfer.RevoluteAxis = RefHelixDir;

            catPart.Update();



            catPart.Update();

            Slot GewindeRille = SF2D.AddNewSlotFromRef(RefmyGewinde, RefHelix);

            Reference RefmyPad = catPart.CreateReferenceFromObject(catSchaftBlock);
            HybridShapeSurfaceExplicit GewindestangenSurface = HSF2D.AddNewSurfaceDatum(RefmyPad);
            Reference RefGewindestangenSurface = catPart.CreateReferenceFromObject(GewindestangenSurface);

            GewindeRille.ReferenceSurfaceElement = RefGewindestangenSurface;

            Reference RefGewindeRille = catPart.CreateReferenceFromObject(GewindeRille);

            catPart.Update();
        }

        //Skizze erzeugen für Helix
        private Sketch makeGewindeSkizzeTrapez(Schraube Schraubeneigenschaft)
        {
            Double Steigung = Schraubeneigenschaft.steigung;
            Double Ri = Schraubeneigenschaft.innenradius;

            OriginElements catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference RefmyPlaneZX = (Reference)catOriginElements.PlaneZX;

            Sketch gewinde = catSketches.Add(RefmyPlaneZX);
            catPart.InWorkObject = gewinde;
            gewinde.set_Name("Gewinde");

            double V_oben_links = -(((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan((30 * Math.PI) / 180));
            double H_oben_links = Ri;

            double V_oben_rechts = (((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan(60 * Math.PI) / 180);
            double H_oben_rechts = Ri;

            double V_unten_links = -((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180));
            double H_unten_links = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            double V_unten_rechts = (0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180);
            double H_unten_rechts = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            double V_Mittelpunkt = 0;
            double H_Mittelpunkt = Ri - ((0.6134 * Steigung) - 0.1443 * Steigung);

            Factory2D catFactory2D2 = gewinde.OpenEdition();

            Point2D Oben_links = catFactory2D2.CreatePoint(H_oben_links, V_oben_links);
            Point2D Oben_rechts = catFactory2D2.CreatePoint(H_oben_rechts, V_oben_rechts);
            Point2D Unten_links = catFactory2D2.CreatePoint(H_unten_links, V_unten_links);
            Point2D Unten_rechts = catFactory2D2.CreatePoint(H_unten_rechts, V_unten_rechts);
            Point2D Mittelpunkt = catFactory2D2.CreatePoint(H_Mittelpunkt, V_Mittelpunkt);

            Line2D Oben = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_oben_rechts, V_oben_rechts);
            Oben.StartPoint = Oben_links;
            Oben.EndPoint = Oben_rechts;

            Line2D Rechts = catFactory2D2.CreateLine(H_oben_rechts, V_oben_rechts, H_unten_rechts, V_unten_rechts);
            Rechts.StartPoint = Oben_rechts;
            Rechts.EndPoint = Unten_rechts;

            Circle2D Verrundung = catFactory2D2.CreateCircle(H_Mittelpunkt, V_Mittelpunkt, 0.1 * Steigung, 0, 0);
            Verrundung.CenterPoint = Mittelpunkt;
            Verrundung.StartPoint = Unten_rechts;
            Verrundung.EndPoint = Unten_links;

            Line2D Links = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_unten_links, V_unten_links);
            Links.StartPoint = Unten_links;
            Links.EndPoint = Oben_links;

            gewinde.CloseEdition();
            catPart.Update();

            return gewinde;
        }

        private Sketch makeGewindeSkizzeWithworth(Schraube Schraubeneigenschaft)
        {
            Double Steigung = Schraubeneigenschaft.steigung;
            Double Ri = Schraubeneigenschaft.innenradius;

            OriginElements catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference RefmyPlaneZX = (Reference)catOriginElements.PlaneZX;

            Sketch gewinde = catSketches.Add(RefmyPlaneZX);
            catPart.InWorkObject = gewinde;
            gewinde.set_Name("Gewinde");

            double V_oben_links = -(((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan((30 * Math.PI) / 180));
            double H_oben_links = Ri;

            double V_oben_rechts = (((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan((30 * Math.PI) / 180));
            double H_oben_rechts = Ri;

            double V_unten_links = -((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180));
            double H_unten_links = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            double V_unten_rechts = (0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180);
            double H_unten_rechts = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            double V_Mittelpunkt = 0;
            double H_Mittelpunkt = Ri - ((0.6134 * Steigung) - 0.1443 * Steigung);

            Factory2D catFactory2D2 = gewinde.OpenEdition();

            Point2D Oben_links = catFactory2D2.CreatePoint(H_oben_links, V_oben_links);
            Point2D Oben_rechts = catFactory2D2.CreatePoint(H_oben_rechts, V_oben_rechts);
            Point2D Unten_links = catFactory2D2.CreatePoint(H_unten_links, V_unten_links);
            Point2D Unten_rechts = catFactory2D2.CreatePoint(H_unten_rechts, V_unten_rechts);
            Point2D Mittelpunkt = catFactory2D2.CreatePoint(H_Mittelpunkt, V_Mittelpunkt);

            Line2D Oben = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_oben_rechts, V_oben_rechts);
            Oben.StartPoint = Oben_links;
            Oben.EndPoint = Oben_rechts;

            Line2D Rechts = catFactory2D2.CreateLine(H_oben_rechts, V_oben_rechts, H_unten_rechts, V_unten_rechts);
            Rechts.StartPoint = Oben_rechts;
            Rechts.EndPoint = Unten_rechts;

            Circle2D Verrundung = catFactory2D2.CreateCircle(H_Mittelpunkt, V_Mittelpunkt, 0.1443 * Steigung, 0, 0);
            Verrundung.CenterPoint = Mittelpunkt;
            Verrundung.StartPoint = Unten_rechts;
            Verrundung.EndPoint = Unten_links;

            Line2D Links = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_unten_links, V_unten_links);
            Links.StartPoint = Unten_links;
            Links.EndPoint = Oben_links;

            gewinde.CloseEdition();
            catPart.Update();

            return gewinde;
        }

        private Sketch makeGewindeSkizzeSägen(Schraube Schraubeneigenschaft)
        {
            Double Steigung = Schraubeneigenschaft.steigung;
            Double Ri = Schraubeneigenschaft.innenradius;

            OriginElements catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference RefmyPlaneZX = (Reference)catOriginElements.PlaneZX;

            Sketch gewinde = catSketches.Add(RefmyPlaneZX);
            catPart.InWorkObject = gewinde;
            gewinde.set_Name("Gewinde");

            double V_oben_links = -(((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan((15 * Math.PI) / 180));
            double H_oben_links = Ri;

            double V_oben_rechts = (((((Math.Sqrt(3) / 2) * Steigung) / 6) + 0.6134 * Steigung) * Math.Tan((15 * Math.PI) / 180));
            double H_oben_rechts = Ri;

            double V_unten_links = -((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180));
            double H_unten_links = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            double V_unten_rechts = (0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180);
            double H_unten_rechts = Ri - (0.6134 * Steigung - 0.1443 * Steigung) - Math.Sqrt(Math.Pow((0.1443 * Steigung), 2) - Math.Pow((0.1443 * Steigung) * Math.Sin((60 * Math.PI) / 180), 2));

            Factory2D catFactory2D2 = gewinde.OpenEdition();

            Point2D Oben_links = catFactory2D2.CreatePoint(H_oben_links, V_oben_links);
            Point2D Oben_rechts = catFactory2D2.CreatePoint(H_oben_rechts, V_oben_rechts);
            Point2D Unten_links = catFactory2D2.CreatePoint(H_unten_links, V_unten_links);
            Point2D Unten_rechts = catFactory2D2.CreatePoint(H_unten_rechts, V_unten_rechts);


            Line2D Oben = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_oben_rechts, V_oben_rechts);
            Oben.StartPoint = Oben_links;
            Oben.EndPoint = Oben_rechts;

            Line2D Rechts = catFactory2D2.CreateLine(H_oben_rechts, V_oben_rechts, H_unten_rechts, V_unten_rechts);
            Rechts.StartPoint = Oben_rechts;
            Rechts.EndPoint = Unten_rechts;

            Line2D Unten = catFactory2D2.CreateLine(H_unten_rechts, V_unten_rechts, H_unten_links, V_unten_links);

            Unten.StartPoint = Unten_rechts;
            Unten.EndPoint = Unten_links;

            Line2D Links = catFactory2D2.CreateLine(H_oben_links, V_oben_links, H_unten_links, V_unten_links);
            Links.StartPoint = Unten_links;
            Links.EndPoint = Oben_links;

            gewinde.CloseEdition();
            catPart.Update();

            return gewinde;
        }
        #endregion
        #endregion



        //KOPF
        //Kopf Skizze erzeugen (gilt für Vier- und Sechkant)
        #region Skizze erzeugen 
        internal void ErzeugeKopfSkizze(double entfernung)
        {
            HybridBodies HybridBodyKopf1 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodyKopf1;

            try
            {
                catHybridBodyKopf1 = HybridBodyKopf1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            catSketches = catHybridBodyKopf1.HybridSketches;
            catOriginElements = stg_catiaPart.Part.OriginElements;

            hybridShapePlaneOffset1 = HSF2D.AddNewPlaneOffset((Reference)catOriginElements.PlaneYZ, entfernung, false);

            Reference catref = (Reference)catOriginElements.PlaneYZ;
            stg_catiaKopfProfil = catSketches.Add(catref);
            catHybridBodyKopf1.AppendHybridShape(hybridShapePlaneOffset1);

            stg_catiaPart.Part.InWorkObject = hybridShapePlaneOffset1;
            hybridShapePlaneOffset1.set_Name("OffsetEbene");
            catPart.Update();

            HybridShapes hybridShapes1 = catHybridBodyKopf1.HybridShapes;
            Reference catReference1 = (Reference)hybridShapes1.Item("OffsetEbene");

            stg_catiaKopfProfil = catSketches.Add(catReference1);

            ErzeugeAchsensystem();
        }
        #endregion

        //Vierkantkopf erzeugen
        #region Vierkantkopf erzeugen 
        internal void ErzeugeVierkantKopfProfil(Schraubenkopf kopfeigenschaften)
        {
            double Kopfhöhe = kopfeigenschaften.höhe;
            double HalbeKopfbreite = (kopfeigenschaften.breite) / 2;


            F2D = stg_catiaKopfProfil.OpenEdition();

            Point2D Point2D_1 = F2D.CreatePoint(HalbeKopfbreite, HalbeKopfbreite);                   //Viereck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(-HalbeKopfbreite, HalbeKopfbreite);
            Point2D Point2D_3 = F2D.CreatePoint(-HalbeKopfbreite, -HalbeKopfbreite);
            Point2D Point2D_4 = F2D.CreatePoint(HalbeKopfbreite, -HalbeKopfbreite);

            Line2D Line2D_1 = F2D.CreateLine(HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;
            Line2D Line2D_2 = F2D.CreateLine(-HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite, -HalbeKopfbreite);
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;
            Line2D Line2D_3 = F2D.CreateLine(-HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite);
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;
            Line2D Line2D_4 = F2D.CreateLine(HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite, HalbeKopfbreite);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_1;

            stg_catiaKopfProfil.CloseEdition();
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;                     //Block erzeugen 
            Pad catKopfBlock = SF2D.AddNewPad(stg_catiaKopfProfil, Kopfhöhe);
            catKopfBlock.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            catKopfBlock.set_Name("Schraubenkopf");

            stg_catiaPart.Part.Update();

            ErzeugeUnterlegscheibe(HalbeKopfbreite, Kopfhöhe);

          
        }
        #endregion

        //Schutzprofil für Vierkantkopf
        #region Schutzprofil
        internal void ErzeugeUnterlegscheibe(double halbekopfbreite, double kopfhöhe)
        {
            double durchmesser = 2 * ((13 / 10) * halbekopfbreite);
            double höhe = kopfhöhe / 10;

            HybridBodies HybridBodySchutz1 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodySchutz1;

            try
            {
                catHybridBodySchutz1 = HybridBodySchutz1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }

            catSketches = catHybridBodySchutz1.HybridSketches;             //Skizze erzeugen
            catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference catref1 = (Reference)hybridShapePlaneOffset1;
            MECMOD.Sketch stg_catiaSchutzProfil = catSketches.Add(catref1);
            ErzeugeAchsensystem();

            F2D = stg_catiaSchutzProfil.OpenEdition();

            Point2D Mittelpunkt1 = F2D.CreatePoint(0, 0);              //Profil erzeugen
            Circle2D Schutz1 = F2D.CreateCircle(0, 0, durchmesser, 0, 0);
            Schutz1.CenterPoint = Mittelpunkt1;

            stg_catiaSchutzProfil.CloseEdition();

            //Block erzeugen
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;            //Block erzeugen
            Pad SchutzBlock1 = SF2D.AddNewPad(stg_catiaSchutzProfil, höhe);
            SchutzBlock1.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            SchutzBlock1.set_Name("Schutzscheibe");

            stg_catiaPart.Part.Update();


            //Fase und Kantenverrundungen
            #region Fase und Kantenverrundungen
            Reference referenz3 = catPart.CreateReferenceFromName("");

            CatFilletEdgePropagation catEdgeProp1 = CatFilletEdgePropagation.catTangencyFilletEdgePropagation;
            ConstRadEdgeFillet kantenverrundung1 = SF2D.AddNewEdgeFilletWithConstantRadius(referenz3, catEdgeProp1 ,0.3 * höhe);

            Reference referenz4 = catPart.CreateReferenceFromBRepName("REdge:(Edge:(Face:(Brp:(Pad.3;0:(Brp:(Sketch.4;1)));None:();Cf11:());Face:(Brp:(Pad.3;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", SchutzBlock1);

            kantenverrundung1.AddObjectToFillet(referenz4);

            Reference referenz5 = catPart.CreateReferenceFromName("");

            CatFilletEdgePropagation catEdgeProp2 = CatFilletEdgePropagation.catTangencyFilletEdgePropagation;
            ConstRadEdgeFillet kantenverrundung2 = SF2D.AddNewEdgeFilletWithConstantRadius(referenz5, catEdgeProp2, 0.3 * höhe);

            Reference referenz6 = catPart.CreateReferenceFromBRepName("REdge:(Edge:(Face:(Brp:((Brp:(Pad.3;1);Brp:(Pad.2;1)));None:();Cf11:());Face:(Brp:(Pad.3;0:(Brp:(Sketch.4;1)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", SchutzBlock1);

            kantenverrundung2.AddObjectToFillet(referenz6);

            stg_catiaPart.Part.Update();

            if (kopfhöhe > 2)
            {

                Reference referenz7 = catPart.CreateReferenceFromName("");

                CatChamferPropagation catChamProp = CatChamferPropagation.catTangencyChamfer;
                CatChamferOrientation CatChamOrien = CatChamferOrientation.catNoReverseChamfer;
                CatChamferMode catChamMode = CatChamferMode.catLengthAngleChamfer;

                Chamfer Fase2 = SF2D.AddNewChamfer(referenz7, catChamProp, catChamMode, CatChamOrien, 1, 45);

                Reference referenz8 = catPart.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.2;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", kantenverrundung2);

                Fase2.AddElementToChamfer(referenz8);

                stg_catiaPart.Part.Update();

            }
            #endregion
        }
        #endregion

        //Sechskantkopf erzeugen
        #region Sechskantkopf
        internal void ErzeugeSechskantKopfProfil(Schraubenkopf kopfeigenschaften)
        {
            double SW = kopfeigenschaften.breite;
            double SWWurzel3 = SW / Math.Sqrt(3);
            double SWWurzel2 = SW / Math.Sqrt(2);
            double Kopfhöhe = kopfeigenschaften.höhe;
            double V0 = 0;

            F2D = stg_catiaKopfProfil.OpenEdition();

            Point2D Point2D_1 = F2D.CreatePoint((SWWurzel2 / 2), (- SW / 2));                              //Sechseck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(SWWurzel3, 0);
            Point2D Point2D_3 = F2D.CreatePoint((SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_4 = F2D.CreatePoint((-SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_5 = F2D.CreatePoint(-SWWurzel3, 0);
            Point2D Point2D_6 = F2D.CreatePoint((-SWWurzel2 /2), (-SW / 2));
            

            Line2D Line2D_1 = F2D.CreateLine((SWWurzel2 / 2), (-SW / 2), SWWurzel3, 0);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = F2D.CreateLine(SWWurzel3, 0, (SWWurzel2 / 2), (SW / 2));
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = F2D.CreateLine((SWWurzel2 / 2), (SW / 2), (-SWWurzel2 / 2), (SW / 2));
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = F2D.CreateLine((-SWWurzel2 / 2), (SW / 2), -SWWurzel3, 0);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = F2D.CreateLine(-SWWurzel3, 0, (-SWWurzel2 / 2), (-SW / 2));
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = F2D.CreateLine((-SWWurzel2 / 2), (-SW / 2), (SWWurzel2 / 2), (-SW / 2));
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_1;

            
            stg_catiaKopfProfil.CloseEdition();                                                //Main Body in Bearbeitung definieren 
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            Pad catKopfBlock = SF2D.AddNewPad(stg_catiaKopfProfil, Kopfhöhe);                    //Sechskantkopf erzeugen 
            catKopfBlock.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            catKopfBlock.set_Name("Innensechskant");

            stg_catiaPart.Part.Update();

        }
        #endregion

        //Innensechskantkopf erzeugen
        #region Innensechskantkopf 
        internal void ErzeugeInnensechskantkopfSkizze(double entfernung)
        {
            HybridBodies HybridBodyKopf1 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodyKopf1;

            try
            {
                catHybridBodyKopf1 = HybridBodyKopf1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            catSketches = catHybridBodyKopf1.HybridSketches;
            catOriginElements = stg_catiaPart.Part.OriginElements;

            hybridShapePlaneOffset1 = HSF2D.AddNewPlaneOffset((Reference)catOriginElements.PlaneYZ, entfernung, false);

            Reference catref = (Reference)catOriginElements.PlaneYZ;
            stg_catiaKopfProfil = catSketches.Add(catref);
            catHybridBodyKopf1.AppendHybridShape(hybridShapePlaneOffset1);

            stg_catiaPart.Part.InWorkObject = hybridShapePlaneOffset1;
            hybridShapePlaneOffset1.set_Name("OffsetEbene");
            catPart.Update();

            HybridShapes hybridShapes1 = catHybridBodyKopf1.HybridShapes;
            Reference catReference1 = (Reference)hybridShapes1.Item("OffsetEbene");

            stg_catiaKopfProfil = catSketches.Add(catReference1);

            ErzeugeAchsensystem();
        }

        //Profil und Block für den Innensechskant wird erzeugt
        #region Profil erzeugen
        internal void ErzeugeInnensechskantkopfProfil(Innensechskantkopf kopfeigenschaften)
        {
            F2D = stg_catiaKopfProfil.OpenEdition();

            double radius = kopfeigenschaften.zylinderdurchmesser / 2;
            double außenhöhe = kopfeigenschaften.höhe;
            double schlüsselweite = kopfeigenschaften.innenschlüsselweite;
            double innenhöhe = kopfeigenschaften.innenhöhe;


            Point2D Mittelpunkt = F2D.CreatePoint(0, 0);
            Circle2D KopfSkizze = F2D.CreateCircle(0, 0, radius, 0, 0);

            stg_catiaKopfProfil.CloseEdition();
            stg_catiaPart.Part.InWorkObject = catBody;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;

            Pad catKopfBlock = SF2D.AddNewPad(stg_catiaKopfProfil, außenhöhe);
            catKopfBlock.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            catKopfBlock.set_Name("SchraubenKopf");

            stg_catiaPart.Part.Update();

        }
        #endregion

        //Tasche für den Innensechskant wird erzeugt
        #region Innensechskant erzeugen
        internal void ErzeugeInnenTasche(Innensechskantkopf itsKopfeigenschaften)
        {
            Reference referenz1 = stg_catiaPart.Part.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Pad.2;2);None:();Cf11:());Pad.2_ResultOUT;Z0;G8251)");

            HybridBodies catHybrodBodiesTasche = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodyTasche;

            catHybridBodyTasche = catHybrodBodiesTasche.Item("Geometrisches Set.1");
            Sketches sketchTasche = (Sketches)catHybridBodyTasche.HybridSketches;

            stg_catiaTasche = sketchTasche.Add(referenz1);

            ErzeugeAchsensystem();

            stg_catiaPart.Part.Update();

            F2D = stg_catiaTasche.OpenEdition();

            double SW = itsKopfeigenschaften.innenschlüsselweite;
            double SWWurzel3 = SW / Math.Sqrt(3);
            double SWWurzel2 = SW / Math.Sqrt(2);
            double Kopfhöhe = itsKopfeigenschaften.innenhöhe;


            Point2D Point2D_1 = F2D.CreatePoint((SWWurzel2 / 2), (-SW / 2));                              //Sechseck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(SWWurzel3, 0);
            Point2D Point2D_3 = F2D.CreatePoint((SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_4 = F2D.CreatePoint((-SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_5 = F2D.CreatePoint(-SWWurzel3, 0);
            Point2D Point2D_6 = F2D.CreatePoint((-SWWurzel2 / 2), (-SW / 2));


            Line2D Line2D_1 = F2D.CreateLine((SWWurzel2 / 2), (-SW / 2), SWWurzel3, 0);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = F2D.CreateLine(SWWurzel3, 0, (SWWurzel2 / 2), (SW / 2));
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = F2D.CreateLine((SWWurzel2 / 2), (SW / 2), (-SWWurzel2 / 2), (SW / 2));
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = F2D.CreateLine((-SWWurzel2 / 2), (SW / 2), -SWWurzel3, 0);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = F2D.CreateLine(-SWWurzel3, 0, (-SWWurzel2 / 2), (-SW / 2));
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = F2D.CreateLine((-SWWurzel2 / 2), (-SW / 2), (SWWurzel2 / 2), (-SW / 2));
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_1;

            stg_catiaTasche.CloseEdition();                                         // Main Body in Bearbeitung definieren 
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            Pocket catKopfTasche = SF2D.AddNewPocket(stg_catiaTasche, Kopfhöhe);                    //Sechskantkopf erzeugen 
            catKopfTasche.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
            catKopfTasche.set_Name("Schraubenkopf");

            stg_catiaPart.Part.Update();

            //Kantenverrundung Kopf
            #region Kantenverrundung Kopf
            if (itsKopfeigenschaften.zylinderdurchmesser >= (itsKopfeigenschaften.innenschlüsselweite/2) + itsKopfeigenschaften.innenschlüsselweite && itsKopfeigenschaften.zylinderdurchmesser >= 2)
            {

                
                Reference refernz1 = catPart.CreateReferenceFromName("");

                CatFilletEdgePropagation catEdgeProp = CatFilletEdgePropagation.catTangencyFilletEdgePropagation;

                ConstRadEdgeFillet kantenverrundung1 = SF2D.AddNewEdgeFilletWithConstantRadius(referenz1, catEdgeProp, 2);

                Reference referenz2 = catPart.CreateReferenceFromBRepName("REdge:(Edge:(Face:(Brp:(Pad.2;0:(Brp:(Sketch.5;1)));None:();Cf11:());Face:(Brp:(Pad.2;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", catKopfTasche);
                kantenverrundung1.AddObjectToFillet(referenz2);

                stg_catiaPart.Part.Update();

                
            }
            
            
            #endregion
        }
        #endregion
        #endregion

        //NORMGERECHTES GEWINDE
        //Skizze für Schaft erzeugen 
        #region Schaft Skizze erzeugen
        internal void NG_ErzeugeSchaftSkizze()
        {
            {
                SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;
                HybridBodies HybridBodySchaft4 = stg_catiaPart.Part.HybridBodies;
                HybridBody catHybridBodySchaft4;


                try
                {
                    catHybridBodySchaft4 = HybridBodySchaft4.Item("Geometrisches Set.1");
                }
                catch (Exception)
                {
                    Console.WriteLine("Fehler");
                    return;
                }


                catSketches = catHybridBodySchaft4.HybridSketches;
                catOriginElements = stg_catiaPart.Part.OriginElements;
                Reference catref = (Reference)catOriginElements.PlaneYZ;
                stg_NG_CatiaSchaftProfil = catSketches.Add(catref);

                NG_ErzeugeAchsensystem();

                stg_catiaPart.Part.Update();

            }
        }
        #endregion

        //Schaft als Block definieren
        #region Schaft Block erzeugen 
        internal void NG_ErzeugeSchaftBlock(Schraube Schraubeneigenschaften)
        {
            F2D = stg_NG_CatiaSchaftProfil.OpenEdition();

            double durchmesser = Schraubeneigenschaften.innenradius;
            double länge = Schraubeneigenschaften.laenge;
            double gewindelänge = Schraubeneigenschaften.gewindeLaenge;

            Point2D Mittelpunkt = F2D.CreatePoint(0, 0);
            Circle2D SchaftProfil = F2D.CreateCircle(0, 0, durchmesser, 0, 0);
            SchaftProfil.CenterPoint = Mittelpunkt;

            stg_NG_CatiaSchaftProfil.CloseEdition();

            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;
            Pad NG_SchaftBlock = SF2D.AddNewPad(stg_NG_CatiaSchaftProfil, länge);

            stg_catiaPart.Part.Update();

            

            stg_catiaPart.Part.Update();

            Reference RefMantelFlaeche = stg_catiaPart.Part.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;2)));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", NG_SchaftBlock);
            Reference RefFrontFlaeche = stg_catiaPart.Part.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", NG_SchaftBlock);

            //Gewinde erzeugen
            PARTITF.Thread Gewinde1 = SF2D.AddNewThreadWithOutRef();
            Gewinde1.LateralFaceElement = RefMantelFlaeche;
            Gewinde1.LimitFaceElement = RefFrontFlaeche;
            Gewinde1.Diameter = durchmesser * 2;
            Gewinde1.Depth = gewindelänge;
            Gewinde1.Side = CatThreadSide.catRightSide;

            Gewinde1.CreateUserStandardDesignTable("Metric_Thick_Pitch", @"C:\Program Files\Dassault Systemes\B28\win_b64\resources\standard\thread\Metric_Thick_Pitch.xml");
            Gewinde1.Diameter = durchmesser * 2;
            Gewinde1.Pitch = 1.250000;

            stg_catiaPart.Part.Update();





        }
        #endregion

        //Kopf Skizze, Porfil und Block erzeugen 
        #region Kopf Skizze und Block erzeugen 
        internal void NG_ErzeugeKopfSkizze(Schraubenkopf NG_kopfeigenschaften, string kopfart)
        {
            //Skizze erzeugen 
            double Kopfhöhe = NG_kopfeigenschaften.höhe;
            double HalbeKopfbreite = (NG_kopfeigenschaften.breite) / 2;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;
            HybridBodies HybridBodySchaft5 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodySchaft4;


            try
            {
                catHybridBodySchaft4 = HybridBodySchaft5.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }


            catSketches = catHybridBodySchaft4.HybridSketches;
            catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference catref = (Reference)catOriginElements.PlaneYZ;
            stg_NG_catiaKopfProfil = catSketches.Add(catref);

            NG_ErzeugeAchsensystem();

            stg_catiaPart.Part.Update();

            F2D = stg_NG_catiaKopfProfil.OpenEdition();

            //Block erzeugen
            #region Vierkantkopf
            if (kopfart == "Vierkant")
            {
                Point2D Point2D_1 = F2D.CreatePoint(HalbeKopfbreite, HalbeKopfbreite);                   //Viereck erzeugen 
                Point2D Point2D_2 = F2D.CreatePoint(-HalbeKopfbreite, HalbeKopfbreite);
                Point2D Point2D_3 = F2D.CreatePoint(-HalbeKopfbreite, -HalbeKopfbreite);
                Point2D Point2D_4 = F2D.CreatePoint(HalbeKopfbreite, -HalbeKopfbreite);

                Line2D Line2D_1 = F2D.CreateLine(HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite);
                Line2D_1.StartPoint = Point2D_1;
                Line2D_1.EndPoint = Point2D_2;
                Line2D Line2D_2 = F2D.CreateLine(-HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite, -HalbeKopfbreite);
                Line2D_2.StartPoint = Point2D_2;
                Line2D_2.EndPoint = Point2D_3;
                Line2D Line2D_3 = F2D.CreateLine(-HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite, -HalbeKopfbreite);
                Line2D_3.StartPoint = Point2D_3;
                Line2D_3.EndPoint = Point2D_4;
                Line2D Line2D_4 = F2D.CreateLine(HalbeKopfbreite, -HalbeKopfbreite, HalbeKopfbreite, HalbeKopfbreite);
                Line2D_4.StartPoint = Point2D_4;
                Line2D_4.EndPoint = Point2D_1;

                stg_NG_catiaKopfProfil.CloseEdition();
                stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

                SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;                     //Block erzeugen 
                Pad catKopfBlock = SF2D.AddNewPad(stg_NG_catiaKopfProfil, Kopfhöhe);
                catKopfBlock.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
                catKopfBlock.set_Name("Schraubenkopf");

                stg_catiaPart.Part.Update();

                NG_ErzeugeUnterlegscheibe(HalbeKopfbreite, Kopfhöhe);
            }
            #endregion

            #region Secchsantkopf
            else if (kopfart == "Sechskant")
            {
                double SW = (NG_kopfeigenschaften.breite);

                F2D = stg_NG_catiaKopfProfil.OpenEdition();

                double SWWurzel3 = SW / Math.Sqrt(3);
                double SWWurzel2 = SW / Math.Sqrt(2);

                Point2D Point2D_1 = F2D.CreatePoint((SWWurzel2 / 2), (-SW / 2));                              //Sechseck erzeugen 
                Point2D Point2D_2 = F2D.CreatePoint(SWWurzel3, 0);
                Point2D Point2D_3 = F2D.CreatePoint((SWWurzel2 / 2), (SW / 2));
                Point2D Point2D_4 = F2D.CreatePoint((-SWWurzel2 / 2), (SW / 2));
                Point2D Point2D_5 = F2D.CreatePoint(-SWWurzel3, 0);
                Point2D Point2D_6 = F2D.CreatePoint((-SWWurzel2 / 2), (-SW / 2));


                Line2D Line2D_1 = F2D.CreateLine((SWWurzel2 / 2), (-SW / 2), SWWurzel3, 0);
                Line2D_1.StartPoint = Point2D_1;
                Line2D_1.EndPoint = Point2D_2;

                Line2D Line2D_2 = F2D.CreateLine(SWWurzel3, 0, (SWWurzel2 / 2), (SW / 2));
                Line2D_2.StartPoint = Point2D_2;
                Line2D_2.EndPoint = Point2D_3;

                Line2D Line2D_3 = F2D.CreateLine((SWWurzel2 / 2), (SW / 2), (-SWWurzel2 / 2), (SW / 2));
                Line2D_3.StartPoint = Point2D_3;
                Line2D_3.EndPoint = Point2D_4;

                Line2D Line2D_4 = F2D.CreateLine((-SWWurzel2 / 2), (SW / 2), -SWWurzel3, 0);
                Line2D_4.StartPoint = Point2D_4;
                Line2D_4.EndPoint = Point2D_5;

                Line2D Line2D_5 = F2D.CreateLine(-SWWurzel3, 0, (-SWWurzel2 / 2), (-SW / 2));
                Line2D_5.StartPoint = Point2D_5;
                Line2D_5.EndPoint = Point2D_6;

                Line2D Line2D_6 = F2D.CreateLine((-SWWurzel2 / 2), (-SW / 2), (SWWurzel2 / 2), (-SW / 2));
                Line2D_6.StartPoint = Point2D_6;
                Line2D_6.EndPoint = Point2D_1;

                stg_NG_catiaKopfProfil.CloseEdition();                                                //Main Body in Bearbeitung definieren 
                stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

                Pad catKopfBlock = SF2D.AddNewPad(stg_NG_catiaKopfProfil, Kopfhöhe);                    //Sechskantkopf erzeugen 
                catKopfBlock.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
                catKopfBlock.set_Name("Innensechskant");

                stg_catiaPart.Part.Update();
            }
            #endregion


        }

        //Nur für Vierkantkopf: Schutzcheibe erzeugen
        private void NG_ErzeugeUnterlegscheibe(double halbekopfbreite, double kopfhöhe)
        {
            double durchmesser = 2 * ((13 / 10) * halbekopfbreite);
            double höhe = kopfhöhe / 10;

            HybridBodies HybridBodySchutz4 = stg_catiaPart.Part.HybridBodies;
            HybridBody catHybridBodySchutz4;

            try
            {
                catHybridBodySchutz4 = HybridBodySchutz4.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }

            catSketches = catHybridBodySchutz4.HybridSketches;
            catOriginElements = stg_catiaPart.Part.OriginElements;
            Reference catref = (Reference)catOriginElements.PlaneYZ;
            MECMOD.Sketch stg_catiaSchutzProfil = catSketches.Add(catref);

            NG_ErzeugeAchsensystem2();

            F2D = stg_catiaSchutzProfil.OpenEdition();

            Point2D Mittelpunkt1 = F2D.CreatePoint(0, 0);              //Profil erzeugen
            Circle2D Schutz1 = F2D.CreateCircle(0, 0, durchmesser, 0, 0);
            Schutz1.CenterPoint = Mittelpunkt1;

            stg_catiaSchutzProfil.CloseEdition();

            //Block erzeugen
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            SF2D = (ShapeFactory)stg_catiaPart.Part.ShapeFactory;            //Block erzeugen
            Pad SchutzBlock1 = SF2D.AddNewPad(stg_catiaSchutzProfil, höhe);
            SchutzBlock1.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
            SchutzBlock1.set_Name("Schutzscheibe");

            stg_catiaPart.Part.Update();
        }

        #endregion


        //Material erzeugen 
        #region Material
        internal void ErzeugeMaterial(string materialmitgabe)
        {
            String sFilePath = @"C:\Program Files\Dassault Systemes\B28\win_b64\startup\materials\German\Catalog.CATMaterial";
            MaterialDocument oMaterial_document = (MaterialDocument)stg_catiaApp.Documents.Open(sFilePath);
            MaterialFamilies cFamilies_list = oMaterial_document.Families;

            string materialangabe = "";

            foreach (MaterialFamily mf in cFamilies_list)
            {
                Console.WriteLine(mf.get_Name());
            }
            #region Material
            switch (materialmitgabe)
            {
                case "Stahl":
                    {
                        materialangabe = "Stahl";
                    }
                    break;
                case "Eisen":
                    {
                        materialangabe = "Eisen";
                    }
                    break;
                case "Messing":
                    {
                        materialangabe = "Messing";
                    }
                    break;
                case "Gelbes Messing":
                    {
                        materialangabe = "Gelbes Messing";
                    }
                    break;
                case "Kupfer":
                    {
                        materialangabe = "Kupfer";
                    }
                    break;
                case "Aluminium":
                    {
                        materialangabe = "Aluminium";
                    }
                    break;
                case "Silber":
                    {
                        materialangabe = "Silber";
                    }
                    break;


            }
            #endregion

            MaterialFamily myMf = cFamilies_list.Item("Metall");
            foreach (Material mat in myMf.Materials)
            {
                Console.WriteLine(mat.get_Name());
            }

            Material myStahl = myMf.Materials.Item(materialangabe);

            MaterialManager partMatManager = stg_catiaPart.Part.GetItem("CATMatManagerVBExt") as MaterialManager;

            // brauchen Sie Stahl im Part?
            short linkMode = 0;
            partMatManager.ApplyMaterialOnPart(stg_catiaPart.Part, myStahl, linkMode);

            // brauchen Sie Stahl im Body?
            linkMode = 1;
            partMatManager.ApplyMaterialOnBody(stg_catiaPart.Part.MainBody, myStahl, linkMode);

            oMaterial_document.Close();
        }
        #endregion

        //Screenshot erzeugen
        #region Screenshot
        public void ErzeugeScreenshot(Schraube Schraube)
        {
            //Dateiname festlegen
            string bildname = "M" + Schraube.innenradius + "x" + Schraube.laenge;

            //Standardhintergrund speichern
            object[] arr1 = new object[3];
            stg_catiaApp.ActiveWindow.ActiveViewer.GetBackgroundColor(arr1);

        
            //Hintergrund auf weiß setzen
            object[] arr2 = new object[] { 1, 1, 1 };
            stg_catiaApp.ActiveWindow.ActiveViewer.PutBackgroundColor(arr2);

            //3D Kompass ausblenden
            stg_catiaApp.StartCommand("CompassDisplayOff");
            stg_catiaApp.ActiveWindow.ActiveViewer.Reframe();

            INFITF.SettingControllers settingControllers1 = stg_catiaApp.SettingControllers;

        //Screenshot wird erstellt und gespeichert
        stg_catiaApp.ActiveWindow.ActiveViewer.CaptureToFile(CatCaptureFormat.catCaptureFormatBMP, "C:\\Windows\\Temp\\" + bildname + ".bmp");

        //3D Kompass einblenden
        stg_catiaApp.StartCommand("CompassDisplayOn");

        //Setzt die Hintergrundfarbe auf Standard zurück
        stg_catiaApp.ActiveWindow.ActiveViewer.PutBackgroundColor(arr1);

        }


        #endregion

        //Normteile erzeugen
        #region Normteile 
        internal void ErzeugeProduct(double gewindedurchmesser, double gewindelänge, double schlüsselweite, double kopfhöhe)
        {
            INFITF.Documents catDocuments1 = stg_catiaApp.Documents;
            productDocuments1 = catDocuments1.Add("Product") as ProductStructureTypeLib.ProductDocument;

            double entfernung = gewindelänge;

            //Part 1 aus Produkt erzeugen: Schraube
            #region Part 1
            //Product 1
            ProductStructureTypeLib.Product product1 = productDocuments1.Product;
            product1.set_PartNumber("Product");
            product1.set_Name("The_Root_Product");

            ProductStructureTypeLib.Products products1 = product1.Products;

            //Part 1
            Product product2 = products1.AddNewComponent("Part", "");
            product2.set_PartNumber("Product.1");

            PartDocument partDokument1 = (PartDocument)catDocuments1.Item("Product.1.CATPart");

            part1 = partDokument1.Part;
            #endregion

            //Part 2 aus Produkt erzeugen: Mutter
            #region Part 2
            //Product 2
            ProductStructureTypeLib.Product product3 = productDocuments1.Product;
            product3.set_PartNumber("Product");
            product3.set_Name("The_Root_Product");

            ProductStructureTypeLib.Products products3 = product3.Products;

            //Part 2
            Product product5 = products3.AddNewComponent("Part", "");
            product5.set_PartNumber("Product.2");

            PartDocument partDokument = (PartDocument)catDocuments1.Item("Product.2.CATPart");

            part2 = partDokument.Part;
            #endregion

            //Part 3 Unterlegscheibe
            #region Part 3
            //Product 3
            ProductStructureTypeLib.Product product4 = productDocuments1.Product;
            product4.set_PartNumber("Product");
            product4.set_Name("The_Root_Product");

            ProductStructureTypeLib.Products products4 = product4.Products;

            //Part 3
            Product product6 = products4.AddNewComponent("Part", "");
            product6.set_PartNumber("Product.3");

            PartDocument partDokument2 = (PartDocument)catDocuments1.Item("Product.3.CATPart");

            part3 = partDokument2.Part;
            #endregion

            //Skizze Part 1: Schraube erzeugen
            #region Part 1 

            HybridBodies HybridBodyKopf1 = part1.HybridBodies;
            HybridBody catHybridBody1;

            try
            {
                catHybridBody1 = HybridBodyKopf1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches1 = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = part1.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;
            stg_cat_Profil1 = catSketches1.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystemX();

            // Part aktualisieren
            part1.Update();

            // Skizze umbenennen
            stg_cat_Profil1.set_Name("Rechteck");

            //Schaft erzeugen
            catFactory2D = stg_cat_Profil1.OpenEdition();

            Circle2D circle = catFactory2D.CreateCircle(0, 0, gewindedurchmesser/ 2, 0, 0);

            // Skizzierer verlassen

            stg_cat_Profil1.CloseEdition();

            // Part aktualisieren

            part1.Update();

            // Hauptkoerper in Bearbeitung definieren
            part1.InWorkObject = part1.MainBody;

            //erzeugen
            ShapeFactory catShapeFactory1 = (ShapeFactory)part1.ShapeFactory;
            Pad catPad1 = catShapeFactory1.AddNewPad(stg_cat_Profil1, gewindelänge);

            // Block umbenennen
            catPad1.set_Name("Schaft");

            Reference RefMantelFlaeche = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;1)));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", catPad1);
            Reference RefFrontFlaeche = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", catPad1);

            //Gewinde erzeugen
            PARTITF.Thread Gewinde1 = catShapeFactory1.AddNewThreadWithOutRef();
            Gewinde1.LateralFaceElement = RefMantelFlaeche;
            Gewinde1.LimitFaceElement = RefFrontFlaeche;
            Gewinde1.Diameter = gewindedurchmesser;
            Gewinde1.Depth = gewindelänge;
            Gewinde1.Side = CatThreadSide.catRightSide;

            Gewinde1.CreateUserStandardDesignTable("Metric_Thick_Pitch", @"C:\Program Files\Dassault Systemes\B28\win_b64\resources\standard\thread\Metric_Thick_Pitch.xml");
            Gewinde1.Diameter = gewindedurchmesser;
            Gewinde1.Pitch = 1.250000;

            part1.Update();

            // Part aktualisieren
            part1.Update();

            part2.Update();
            #endregion

            //Kopf für die Schraube definieren
            ErzeugeKopf(entfernung, schlüsselweite, kopfhöhe);

        }

        #region Kopf erzeugen
        private void ErzeugeKopf(double entfernung, double schlüsselweite, double kopfhöhe)
        {
            HybridShapeFactory HSF = (HybridShapeFactory) part1.HybridShapeFactory;
            HybridBodies HybridBodyKopf5 = part1.HybridBodies;
            HybridBody catHybridBodyKopf5;

            try
            {
                catHybridBodyKopf5 = HybridBodyKopf5.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            catSketches = catHybridBodyKopf5.HybridSketches;
            catOriginElements = part1.OriginElements;

            hybridShapePlaneOffset1 = HSF.AddNewPlaneOffset((Reference)catOriginElements.PlaneYZ, entfernung, false);

            OriginElements origin = part1.OriginElements;
            Reference catref = (Reference)origin.PlaneYZ;
            stg_cat_KopfProfilx = catSketches.Add(catref);
            catHybridBodyKopf5.AppendHybridShape(hybridShapePlaneOffset1);

            part1.InWorkObject = hybridShapePlaneOffset1;
            hybridShapePlaneOffset1.set_Name("OffsetEbene");
            part1.Update();

            HybridShapes hybridShapes1 = catHybridBodyKopf5.HybridShapes;
            Reference catReference1 = (Reference)hybridShapes1.Item("OffsetEbene");

            stg_cat_KopfProfilx = catSketches.Add(catReference1);

            ErzeugeAchsensystemS();

            double SW = schlüsselweite;
            double SWWurzel3 = SW / Math.Sqrt(3);
            double SWWurzel2 = SW / Math.Sqrt(2);
            double Kopfhöhe = kopfhöhe;
            

            F2D = stg_cat_KopfProfilx.OpenEdition();

            Point2D Point2D_1 = F2D.CreatePoint((SWWurzel2 / 2), (-SW / 2));                              //Sechseck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(SWWurzel3, 0);
            Point2D Point2D_3 = F2D.CreatePoint((SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_4 = F2D.CreatePoint((-SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_5 = F2D.CreatePoint(-SWWurzel3, 0);
            Point2D Point2D_6 = F2D.CreatePoint((-SWWurzel2 / 2), (-SW / 2));


            Line2D Line2D_1 = F2D.CreateLine((SWWurzel2 / 2), (-SW / 2), SWWurzel3, 0);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = F2D.CreateLine(SWWurzel3, 0, (SWWurzel2 / 2), (SW / 2));
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = F2D.CreateLine((SWWurzel2 / 2), (SW / 2), (-SWWurzel2 / 2), (SW / 2));
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = F2D.CreateLine((-SWWurzel2 / 2), (SW / 2), -SWWurzel3, 0);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = F2D.CreateLine(-SWWurzel3, 0, (-SWWurzel2 / 2), (-SW / 2));
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = F2D.CreateLine((-SWWurzel2 / 2), (-SW / 2), (SWWurzel2 / 2), (-SW / 2));
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_1;


            stg_cat_KopfProfilx.CloseEdition();                                                //Main Body in Bearbeitung definieren 
            part1.InWorkObject = part1.MainBody;

            ShapeFactory shapefac = (ShapeFactory) part1.ShapeFactory;
            Pad catKopfBlockx = shapefac.AddNewPad(stg_cat_KopfProfilx, Kopfhöhe);                    //Sechskantkopf erzeugen 
            catKopfBlockx.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            catKopfBlockx.set_Name("Sechskant");

            part1.Update();

        }
        #endregion

        //Mutter als Part im Produkt erzeugen
        #region Mutter erzeugen
        internal void ErzeugeProductTeil2(double mutterbreite, double mutterhöhe, double mutterdurchmesser)
        {
            //Normteil Mutter
            double breite = mutterbreite * (5 / 3);

            HybridBodies HybridBodyKopf2 = part2.HybridBodies;
            HybridBody catHybridBody2;

            try
            {
                catHybridBody2 = HybridBodyKopf2.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches2 = catHybridBody2.HybridSketches;
            OriginElements catOriginElements2 = part2.OriginElements;
            Reference catReference2 = (Reference)catOriginElements2.PlaneYZ;
            stg_cat_Profil2 = catSketches2.Add(catReference2);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystemY();

            // Part aktualisieren
            part2.Update();

            catFac = stg_cat_Profil2.OpenEdition();

            //Mutter Skizze erzeugen

            double SW = breite;
            double SWWurzel3 = SW / Math.Sqrt(3);
            double SWWurzel2 = SW / Math.Sqrt(2);
            
            

            Point2D Point2D_1 = catFac.CreatePoint((SWWurzel2 / 2), (-SW / 2));                              //Sechseck erzeugen 
            Point2D Point2D_2 = catFac.CreatePoint(SWWurzel3, 0);
            Point2D Point2D_3 = catFac.CreatePoint((SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_4 = catFac.CreatePoint((-SWWurzel2 / 2), (SW / 2));
            Point2D Point2D_5 = catFac.CreatePoint(-SWWurzel3, 0);
            Point2D Point2D_6 = catFac.CreatePoint((-SWWurzel2 / 2), (-SW / 2));


            Line2D Line2D_1 = catFac.CreateLine((SWWurzel2 / 2), (-SW / 2), SWWurzel3, 0);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = catFac.CreateLine(SWWurzel3, 0, (SWWurzel2 / 2), (SW / 2));
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = catFac.CreateLine((SWWurzel2 / 2), (SW / 2), (-SWWurzel2 / 2), (SW / 2));
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = catFac.CreateLine((-SWWurzel2 / 2), (SW / 2), -SWWurzel3, 0);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = catFac.CreateLine(-SWWurzel3, 0, (-SWWurzel2 / 2), (-SW / 2));
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = catFac.CreateLine((-SWWurzel2 / 2), (-SW / 2), (SWWurzel2 / 2), (-SW / 2));
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_1;

            stg_cat_Profil2.CloseEdition();

            part2.Update();

            part2.InWorkObject = part2.MainBody;

            ShapeFactory catShapeFactory2 = (ShapeFactory)part2.ShapeFactory;
            Pad catPad2 = catShapeFactory2.AddNewPad(stg_cat_Profil2, mutterhöhe);
            catPad2.DirectionOrientation = CatPrismOrientation.catInverseOrientation;

            // Block umbenennen
            catPad2.set_Name("Mutter");

            // Part aktualisieren
            part2.Update();

            ErzeugeMutterBohrung(mutterdurchmesser, mutterhöhe);

            part1.Update();
        }

        #region Mutter Bohrung
        private void ErzeugeMutterBohrung(double mutterdurchmesser, double mutterhöhe)
        {
            HybridBodies HybridBodyKopf3 = part2.HybridBodies;
            HybridBody catHybridBody3;

            try
            {
                catHybridBody3 = HybridBodyKopf3.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches3 = catHybridBody3.HybridSketches;
            OriginElements catOriginElements3 = part2.OriginElements;
            Reference catReference3 = (Reference)catOriginElements3.PlaneYZ;
            stg_cat_ProfilMutterBohrung = catSketches3.Add(catReference3);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystemY();

            Factory2D catFaca = stg_cat_ProfilMutterBohrung.OpenEdition();

            Point2D Mittelpunkt = catFaca.CreatePoint(0, 0);
            Circle2D KopfSkizze = catFaca.CreateCircle(0, 0, mutterdurchmesser/2, 0, 0);

            stg_cat_ProfilMutterBohrung.CloseEdition();
            part2.InWorkObject = part2.MainBody;

            SF2D = (ShapeFactory)part2.ShapeFactory;

            Pocket catMutterBohrung = SF2D.AddNewPocket(stg_cat_ProfilMutterBohrung, mutterhöhe);
            catMutterBohrung.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
            catMutterBohrung.set_Name("Mutter Bohrung");

            part2.Update();
        }

        #endregion

        //Unterlegscheibe erzeugen
        #region Unterlegscheibe erzeugen
        //Normteil Mutter
        internal void ErzeugeProductTeil3(double gewindedurchmesser, double höhe)
        {
            HybridBodies HybridBodyKopf3 = part3.HybridBodies;
            HybridBody catHybridBody3;
            double Höhe = höhe;

            try
            {
                catHybridBody3 = HybridBodyKopf3.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches3 = catHybridBody3.HybridSketches;
            OriginElements catOriginElements4 = part3.OriginElements;
            Reference catReference1 = (Reference)catOriginElements4.PlaneYZ;
            stg_cat_UnterlegProfil = catSketches3.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystemX();

            // Part aktualisieren
            part3.Update();

            //Schaft erzeugen
            Factory2D catFac3 = stg_cat_UnterlegProfil.OpenEdition();

            Circle2D circle = catFac3.CreateCircle(0, 0, gewindedurchmesser * 2, 0, 0);

            // Skizzierer verlassen

            stg_cat_UnterlegProfil.CloseEdition();

            // Part aktualisieren

            part3.Update();

            // Hauptkoerper in Bearbeitung definieren
            part3.InWorkObject = part3.MainBody;

            //erzeugen
            ShapeFactory catShapeFactory3 = (ShapeFactory)part3.ShapeFactory;
            Pad catPad3 = catShapeFactory3.AddNewPad(stg_cat_UnterlegProfil, 1);
            catPad3.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            // Block umbenennen
            catPad3.set_Name("Unterlegscheibe");

            ErzeugeUnterlegBohrung(gewindedurchmesser, Höhe);

            part1.Update();
            part2.Update();
            part3.Update();
        }

        #region Unterlegscheibe Bohrung
        private void ErzeugeUnterlegBohrung(double gewindedurchmesser, double Höhe)
        {
            HybridBodies HybridBodyKopf5 = part3.HybridBodies;
            HybridBody catHybridBody5;

            try
            {
                catHybridBody5 = HybridBodyKopf5.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                Console.WriteLine("Fehler");
                return;
            }
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches5 = catHybridBody5.HybridSketches;
            OriginElements catOriginElements5 = part3.OriginElements;
            Reference catReference5 = (Reference)catOriginElements5.PlaneYZ;
            stg_cat_UnterlegBohrung = catSketches5.Add(catReference5);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystemZ();

            Factory2D catFaca2 = stg_cat_UnterlegBohrung.OpenEdition();

            Circle2D KopfSkizze = catFaca2.CreateCircle(0, 0, gewindedurchmesser / 2, 0, 0);

            stg_cat_UnterlegBohrung.CloseEdition();
            part3.InWorkObject = part3.MainBody;

            ShapeFactory catShape3 = (ShapeFactory)part3.ShapeFactory;

            Pocket catUnterlegBohrung = catShape3.AddNewPocket(stg_cat_UnterlegBohrung, 1);
            catUnterlegBohrung.DirectionOrientation = CatPrismOrientation.catRegularOrientation;
            catUnterlegBohrung.set_Name("Unterlegscheibe Bohrung");

            part3.Update();
        }
        #endregion

        #endregion



        #endregion





    }








}







#endregion