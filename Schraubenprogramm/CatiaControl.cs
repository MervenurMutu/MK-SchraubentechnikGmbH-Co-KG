using System;
using System.Windows;
using INFITF;
using MECMOD;
using PARTITF;
using HybridShapeTypeLib;
using CATMat;

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

        }
        #endregion

        //Sechskantkopf erzeugen
        #region Sechskantkopf
        internal void ErzeugeSechskantKopfProfil(Schraubenkopf kopfeigenschaften)
        {
            double HalbeLänge = (kopfeigenschaften.breite) / 2;
            double HalbeBreite = HalbeLänge * 0.4;
            double Kopfhöhe = kopfeigenschaften.höhe;

            F2D = stg_catiaKopfProfil.OpenEdition();

            Point2D Point2D_1 = F2D.CreatePoint(HalbeBreite, HalbeLänge);                              //Sechseck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(-HalbeBreite, HalbeLänge);
            Point2D Point2D_3 = F2D.CreatePoint(-HalbeLänge, HalbeBreite);
            Point2D Point2D_4 = F2D.CreatePoint(-HalbeLänge, -HalbeBreite);
            Point2D Point2D_5 = F2D.CreatePoint(-HalbeBreite, -HalbeLänge);
            Point2D Point2D_6 = F2D.CreatePoint(HalbeBreite, -HalbeLänge);
            Point2D Point2D_7 = F2D.CreatePoint(HalbeLänge, -HalbeBreite);
            Point2D Point2D_8 = F2D.CreatePoint(HalbeLänge, HalbeBreite);

            Line2D Line2D_1 = F2D.CreateLine(HalbeBreite, HalbeLänge, -HalbeBreite, HalbeLänge);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = F2D.CreateLine(-HalbeBreite, HalbeLänge, -HalbeLänge, HalbeBreite);
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = F2D.CreateLine(-HalbeLänge, HalbeBreite, -HalbeLänge, -HalbeBreite);
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = F2D.CreateLine(-HalbeLänge, -HalbeBreite, -HalbeBreite, -HalbeLänge);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = F2D.CreateLine(-HalbeBreite, -HalbeLänge, HalbeBreite, -HalbeLänge);
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = F2D.CreateLine(HalbeBreite, -HalbeLänge, HalbeLänge, -HalbeBreite);
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_7;

            Line2D Line2D_7 = F2D.CreateLine(HalbeLänge, -HalbeBreite, HalbeLänge, HalbeBreite);
            Line2D_7.StartPoint = Point2D_7;
            Line2D_7.EndPoint = Point2D_8;

            Line2D Line2D_8 = F2D.CreateLine(HalbeLänge, HalbeBreite, HalbeBreite, HalbeLänge);
            Line2D_8.StartPoint = Point2D_8;
            Line2D_8.EndPoint = Point2D_1;

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

            double HalbeLänge = itsKopfeigenschaften.innenschlüsselweite / 2;
            double HalbeBreite = HalbeLänge * 0.4;
            double Kopfhöhe = itsKopfeigenschaften.innenhöhe;

            Point2D Point2D_1 = F2D.CreatePoint(HalbeBreite, HalbeLänge);                              //Sechseck erzeugen 
            Point2D Point2D_2 = F2D.CreatePoint(-HalbeBreite, HalbeLänge);
            Point2D Point2D_3 = F2D.CreatePoint(-HalbeLänge, HalbeBreite);
            Point2D Point2D_4 = F2D.CreatePoint(-HalbeLänge, -HalbeBreite);
            Point2D Point2D_5 = F2D.CreatePoint(-HalbeBreite, -HalbeLänge);
            Point2D Point2D_6 = F2D.CreatePoint(HalbeBreite, -HalbeLänge);
            Point2D Point2D_7 = F2D.CreatePoint(HalbeLänge, -HalbeBreite);
            Point2D Point2D_8 = F2D.CreatePoint(HalbeLänge, HalbeBreite);

            Line2D Line2D_1 = F2D.CreateLine(HalbeBreite, HalbeLänge, -HalbeBreite, HalbeLänge);
            Line2D_1.StartPoint = Point2D_1;
            Line2D_1.EndPoint = Point2D_2;

            Line2D Line2D_2 = F2D.CreateLine(-HalbeBreite, HalbeLänge, -HalbeLänge, HalbeBreite);
            Line2D_2.StartPoint = Point2D_2;
            Line2D_2.EndPoint = Point2D_3;

            Line2D Line2D_3 = F2D.CreateLine(-HalbeLänge, HalbeBreite, -HalbeLänge, -HalbeBreite);
            Line2D_3.StartPoint = Point2D_3;
            Line2D_3.EndPoint = Point2D_4;

            Line2D Line2D_4 = F2D.CreateLine(-HalbeLänge, -HalbeBreite, -HalbeBreite, -HalbeLänge);
            Line2D_4.StartPoint = Point2D_4;
            Line2D_4.EndPoint = Point2D_5;

            Line2D Line2D_5 = F2D.CreateLine(-HalbeBreite, -HalbeLänge, HalbeBreite, -HalbeLänge);
            Line2D_5.StartPoint = Point2D_5;
            Line2D_5.EndPoint = Point2D_6;

            Line2D Line2D_6 = F2D.CreateLine(HalbeBreite, -HalbeLänge, HalbeLänge, -HalbeBreite);
            Line2D_6.StartPoint = Point2D_6;
            Line2D_6.EndPoint = Point2D_7;

            Line2D Line2D_7 = F2D.CreateLine(HalbeLänge, -HalbeBreite, HalbeLänge, HalbeBreite);
            Line2D_7.StartPoint = Point2D_7;
            Line2D_7.EndPoint = Point2D_8;

            Line2D Line2D_8 = F2D.CreateLine(HalbeLänge, HalbeBreite, HalbeBreite, HalbeLänge);
            Line2D_8.StartPoint = Point2D_8;
            Line2D_8.EndPoint = Point2D_1;

            stg_catiaTasche.CloseEdition();                                         // Main Body in Bearbeitung definieren 
            stg_catiaPart.Part.InWorkObject = stg_catiaPart.Part.MainBody;

            Pocket catKopfTasche = SF2D.AddNewPocket(stg_catiaTasche, Kopfhöhe);                    //Sechskantkopf erzeugen 
            catKopfTasche.DirectionOrientation = CatPrismOrientation.catInverseOrientation;
            catKopfTasche.set_Name("Schraubenkopf");

            stg_catiaPart.Part.Update();
        }
        #endregion
        #endregion



        

    }








}







