using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Reflection;
using RGiesecke.DllExport;
using MinitabAddinTLB;


namespace MyMenu
{
    public static class Globals
    {
        public const string GUID = "FE26C977-3422-49B4-B447-6CB99B01C06D";
        public const string Namespace = "MyMenu";
        public const string NamespaceAndClass = "MyMenu.AddIn";
    }

    [type: ComVisible(true)]
    [type: GuidAttribute(Globals.GUID)]
    [type: ClassInterface(ClassInterfaceType.None)]
    [type: ProgId(Globals.NamespaceAndClass)]

    public class AddIn : IMinitabAddin
    {
        internal static Mtb.Application gMtbApp;

        [method: DllExportAttribute("DllRegisterServer", CallingConvention = CallingConvention.StdCall)]
        public static Int32 DllRegisterServer()
        {
            try
            {
                RegistryKey key = Registry.ClassesRoot.CreateSubKey(Globals.NamespaceAndClass);
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.ClassesRoot.CreateSubKey(Globals.NamespaceAndClass + @"\CLSID");
                key.SetValue("", "{" + Globals.GUID + "}");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + "}");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32");
                key.SetValue("", "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32\1.0.0.0");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\ProdId");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\" + Globals.NamespaceAndClass);
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\" + Globals.NamespaceAndClass + @"\CLSID");
                key.SetValue("", "{" + Globals.GUID + "}");

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + "}");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories");

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32");
                key.SetValue("", "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32\1.0.0.0");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\ProgId");
                key.SetValue("", Globals.NamespaceAndClass);
            }
            catch (Exception ex)
            {

                // Probably didn't have permissions to modify the registry
            }

            return 0;
        }

        public AddIn()
        {
        }

        ~AddIn()
        {
        }

        public void OnConnect(Int32 iHwnd, Object pApp, ref Int32 iFlags)
        {
            // This method is called as Minitab is initializing your add-in.
            // The “iHwnd” parameter is the handle to the main Minitab window.
            // The “pApp” parameter is a reference to the “Minitab Automation object.”
            // You can hold onto either of these for use in your add-in.
            // “iFlags” is used to tell Minitab if your add-in has dynamic menus (i.e. should be reloaded each time
            // Minitab starts up).  Set Flags to 1 for dynamic menus and 0 for static.
            AddIn.gMtbApp = pApp as Mtb.Application;
            // This forces Minitab to retain all commands (even those run by the interactive user):
            AddIn.gMtbApp.Options.SaveCommands = true;
            // Static menus:
            iFlags = 0;
            return;
        }

        public void OnDisconnect()
        {
            // This method is called as Minitab is closing your add-in.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            try
            {
                Marshal.FinalReleaseComObject(AddIn.gMtbApp);
                AddIn.gMtbApp = null;
            }
            catch
            {
            }
            return;
        }

        public String GetName()
        {
            // This method returns the friendly name of your add-in:
            // Both the name and the description of the add-in are stored in the registry.
            return "Example C♯ Minitab Add-In";
        }

        public String GetDescription()
        {
            // This method returns the description of your add-in:
            return "An example Minitab add-in written in C♯ using the “My Menu” functionality.";
        }

        public void GetMenuItems(ref String sMainMenu, ref Array saMenuItems, ref Int32 iFlags)
        {
            // This method returns the text for the main menu and each menu item.
            // You can return "|" to create a menu separator in your menu items.
            sMainMenu = "&My Menu";  // This string is the name of the menu.

            saMenuItems = new String[1];  // The strings in this array are the names of the items on the aforementioned menu.

            saMenuItems.SetValue("Example Test", 0);

            // Flags is not currently used:
            iFlags = 0;

            return;
        }

        public String OnDispatchCommand(Int32 iMenu)
        {
            // This method is called whenever a user selects one of your menu items.
            // The iMenu variable should be equivalent to the menu item index set in “GetMenuItems.”
            String command = String.Empty;
            switch (iMenu)
            {
                case 0:
                    command = "rand 20 c1";
                    break;
                default:
                    break;
            }
            return command;
        }

        public void OnNotify(AddinNotifyType eAddinNotifyType)
        {
            // This method is called when Minitab notifies your add-in that something has changed.
            // Use the “eAddinNotifyType” parameter to figure out what changed.
            // Minitab currently fires no events, so this method is not called.
        }

        public Boolean QueryCustomCommand(String sCommand)
        {
            // This method is called when Minitab asks your Addin if it supports a custom command.
            return false;
        }

        public void ExecuteCustomCommand(String sCommand, ref Array saArgs)
        {
            // This method is called when Minitab asks your add-in to execute a custom command.
        }

    }
}
