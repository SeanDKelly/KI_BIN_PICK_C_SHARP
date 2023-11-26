using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Reflection;

namespace KI_BIN_PICK_C_;

public class Part
{

    // Lokale Variablen definieren
    private SldWorks _sw;
    private ModelDoc2 _part;

    public Part(SldWorks solidWorks, string directory)
    {
        int longstatus = 0;
        int longwarnings = 0;
        // Übergebene Solidworks instanz lokaler Variable zuweisen
        _sw = solidWorks;
        // Bautel im übergebenen Pfad öffnen und lokaler Variable zuweisen
        _part = _sw.OpenDoc6(directory,
                            (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                            "",
                            ref longstatus,
                            ref longwarnings);
        if (_part != null)
        {
            // Bauteil erfolgreich geöffnet
            Console.WriteLine("Bauteil erfolgreich geöffnet: " + _part.GetTitle());
        }

        else
        {
            Console.WriteLine("Fehler beim Öffnen des Bauteils!");
        }
    }
    public void Faces()
    { 
        // Schnittstellen für das Bauteil abrufen
        PartDoc partDoc = (PartDoc)this._part;
        if (partDoc == null)
        {
            Console.WriteLine("Fehler beim Abrufen der Bauteilschnittstelle.");
            return;
        }
    
        // Flächen des Bauteils abrufen
        object[] faces = (object[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);

        // Informationen zu jeder Fläche ausgeben
        Console.WriteLine($"Das Bauteil hat {faces.Length} Flächen.");
        for (int i = 0; i < faces.Length; i++)
        {
            Face2 face = (Face2)faces[i];

            string facetyp = face.GetType().ToString();

            Console.WriteLine("Fläche ist vom typ:", facetyp);

        }
    }
    public double[] CenterOfMass()
    {
        MassProperty2 MyMassProp = null;
        ModelDocExtension Extn = null;
        double[] pmoi;
        double[] vValue;
        double[] value = new double[3];
        double val;


        // Extended Properties of Part
        Extn = _part.Extension;

        // Create mass properties and override options
        MyMassProp = (MassProperty2)Extn.CreateMassProperty2();

        // Use document property units (MKS)
        MyMassProp.UseSystemUnits = false;

        // Schwerpunkt
        value = (double[])MyMassProp.CenterOfMass;

        /*
        Console.WriteLine("Mass properties before override");
        Console.WriteLine("");
        val = MyMassProp.Mass;
        Console.WriteLine("Mass: " + val);
        value = (double[])MyMassProp.CenterOfMass;
        Console.WriteLine("Center of Mass: X:" + value[0] + " Y:" + value[1] + " Z:" + value[2]);
        val = MyMassProp.Volume;
        Console.WriteLine("Volume: " + val);
        val = MyMassProp.Density;
        Console.WriteLine("Density: " + val);
        val = MyMassProp.SurfaceArea;
        Console.WriteLine("Surface area: " + val);
        pmoi = (double[])MyMassProp.PrincipalMomentsOfInertia;
        Console.WriteLine("Principal moments of inertiae: Px: " + pmoi[0] + ", Py: " + pmoi[1] + ", and Pz: " + pmoi[2]);
        vValue = (double[])MyMassProp.GetMomentOfInertia(0);
        Console.WriteLine("Moments of inertia: Lxx: " + vValue[0] + ", Lxy: " + vValue[1] + ", Lxz: " + vValue[2] + ", Lyx: " + vValue[3] + ", Lyy: " + vValue[4] + ", Lyz: " + vValue[5] + ", Lzx: " + vValue[6] + ", Lzy: " + vValue[7] + ", Lzz: " + vValue[8]);
        */

        return value;
    }
}
