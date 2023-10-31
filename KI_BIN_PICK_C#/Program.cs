using System;
using System.Diagnostics;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;





public class SW
{

    public SldWorks swApp = null;


    // Constructor
    public SW()
    {
        // Erstellt eine SolidWorks instanz bzw öffnet SolidWorks

        swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
        swApp.Visible = true; // SolidWorks sichtbar machen
    }

    // Destructor
    ~SW()
    {
        // Ressourcen freigeben bzw SowlidWorks schließen
        /*
        if (swModel != null)
        {
            swModel.Close();
            swModel = null;
        }
        */
        if (swApp != null)
        {
            swApp.ExitApp();
            swApp = null;
        }
       
        Console.WriteLine("Drücken Sie eine beliebige Taste, um das Programm zu beenden.");
        Console.ReadKey();
    }

    // Bauteil öffnen
    public ModelDoc2 OpenPart(string directory)
    {
        int longstatus = 0;
        int longwarnings = 0;
        ModelDoc2 swModel = null;
        directory = "C:\\Users\\domi9\\Documents\\_FH Kempten\\6. Semester\\Projekt MT6\\Bauteile\\UniversalgreifermitKrümmung.SLDPRT";
        //C:\\Users\\domi9\\Documents\\_FH Kempten\\6. Semester\\Projekt MT6\\Bauteile\\UniversalgreifermitKrümmung.SLDPRT
        // "C:\\Users\\theja\\Downloads\\Greifer.SLDPRT"
        swModel = swApp.OpenDoc6(directory,
                                (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                                "",
                                ref longstatus,
                                ref longwarnings);
        if (swModel != null)
        {
            // Bauteil erfolgreich geöffnet
            Console.WriteLine("Bauteil erfolgreich geöffnet: " + swModel.GetTitle());
        }

        else
        {
            Console.WriteLine("Fehler beim Öffnen des Bauteils!");
        }

        return swModel;
    }

    // Schwerpunkt berechnen
    public double[] CenterOfMassPart(ModelDoc2 Part)
    {
        MassProperty2 MyMassProp = null;
        ModelDocExtension Extn = null;
        double[] pmoi;
        double[] vValue;
        double[] value = new double[3];
        double val;


        // Extended Properties of Part
        Extn = Part.Extension;

        // Create mass properties and override options
        MyMassProp = (MassProperty2)Extn.CreateMassProperty2();

        // Use document property units (MKS)
        MyMassProp.UseSystemUnits = false;

        // Schwerpunkt
        value = (double[])MyMassProp.CenterOfMass;

        return value;
    }

}

public class Part
{
    public Part()
    {
        // Erstellen Sie eine Instanz von SOLIDWORKS
        swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
        if (swApp == null)
        {
            throw new Exception("SOLIDWORKS not found.");
        }
    }
    
    public double GetMass()
    {
        if (swApp == null)
        {
            throw new Exception("SOLIDWORKS not available.");
        }

        // Holen Sie das aktuell geöffnete Modell
        ModelDoc2 swModel = swApp.IActiveDoc2;

        if (swModel == null)
        {
            throw new Exception("No active document.");
        }

        // Überprüfen Sie, ob es sich um ein Bauteil handelt
        if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
        {
            throw new Exception("The active document is not a part document.");
        }

        // Ermitteln Sie die Masse des Bauteils
        //double mass;
        double density;
        double mass = 0;

        //int numProperties = (int)swModel.GetMassProperties2(ref mass);

        double[] numProperties = (int)swModel.GetMassProperties2(ref mass);

        //out mass, out density, out _, out _, out _
        
        if (numProperties[0] == 1) // 1 bedeutet, dass die Masse erfolgreich abgerufen wurde
        {
            return mass;
        }

        else
        {
            throw new Exception("Failed to get mass properties.");
        }
    }

    
    public double[] GetCenterOfMassPart(ModelDoc2 Part)
    {
        MassProperty2 MyMassProp = null;
        ModelDocExtension Extn = null;
        double[] pmoi;
        double[] vValue;
        double[] value = new double[3];
        double val;


        // Extended Properties of Part
        Extn = Part.Extension;

        // Create mass properties and override options
        MyMassProp = (MassProperty2)Extn.CreateMassProperty2();

        // Use document property units (MKS)
        MyMassProp.UseSystemUnits = false;

        // Schwerpunkt
        value = (double[])MyMassProp.CenterOfMass;

        return value;
    }


    public SldWorks swApp = null;

    //public ModelDoc2 Part = null;

    //public float Density;

    //double mass;

    public float CenterOfMass;

    //Surface[] HandlingSurface = {};


}

class Program
{
    static void Main()
    {
        Console.WriteLine("Starte SolidWorks");
        SW SWProgramm = new SW();
        Console.WriteLine("Bauteil Öffnen");
        ModelDoc2 Greifer = SWProgramm.OpenPart("C:\\Users\\domi9\\Documents\\_FH Kempten\\6. Semester\\Projekt MT6\\Bauteile\\UniversalgreifermitKrümmung.SLDPRT");
        Console.WriteLine("Schwerpunkt berechnen");
        double[] SchwerpunktVektor = SWProgramm.CenterOfMassPart(Greifer);
        Console.WriteLine("Schwerpunkt: X:" + SchwerpunktVektor[0] + " Y:" + SchwerpunktVektor[1] + " Z:" + SchwerpunktVektor[2]);
        Part Part = new Part();
        double mass = Part.GetMass();
        Console.WriteLine("Masse des geöffneten Bauteils: " + mass);
        double[] SchwerpunktVektor2 = Part.GetCenterOfMassPart(Greifer);
        Console.WriteLine("Schwerpunkt X:" + SchwerpunktVektor2[0] + "Y:" + SchwerpunktVektor2[1] + "Z:" + SchwerpunktVektor2[2]);
    }
}

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
