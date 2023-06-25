using System;
using System.Diagnostics;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;



// SolidWorks-Anwendung erstellen
SldWorks swApp = null;
ModelDoc2 swModel = null;
MassProperty2 MyMassProp = null;
ModelDocExtension Extn = null;
int longstatus = 0;
int longwarnings = 0;
double[] pmoi;
double[] vValue;
double[] value = new double[3];
int errors;
int warnings;
double val;

try
{
    swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
    swApp.Visible = true; // SolidWorks sichtbar machen

    // Bauteil öffnen
    //(int)swOpenDocOptions_e.swOpenDocOptions_ViewOnly,

    swModel = swApp.OpenDoc6("C:\\Users\\theja\\Downloads\\Greifer.SLDPRT",
                                            (int)swDocumentTypes_e.swDocPART,
                                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                                            "",
                                            ref longstatus,
                                            ref longwarnings);
                



    if (swModel != null)
    {
        // Bauteil erfolgreich geöffnet
        Console.WriteLine("Bauteil erfolgreich geöffnet: " + swModel.GetTitle());

        // Schwerpunkt berechnen
        Extn = swModel.Extension;



        // Create mass properties and override options

        MyMassProp = (MassProperty2)Extn.CreateMassProperty2();

        // Use document property units (MKS)

        MyMassProp.UseSystemUnits = false;

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
    }
    else
    {
        Console.WriteLine("Fehler beim Öffnen des Bauteils!");
    }
}
catch (Exception ex)
{
    Console.WriteLine("Fehler: " + ex.Message);
}
finally
{
    // Ressourcen freigeben
    if (swModel != null)
    {
        swModel.Close();
        swModel = null;
    }
    if (swApp != null)
    {
        swApp.ExitApp();
        swApp = null;
    }
}

Console.WriteLine("Drücken Sie eine beliebige Taste, um das Programm zu beenden.");
Console.ReadKey();

