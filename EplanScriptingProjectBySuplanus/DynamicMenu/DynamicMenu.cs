/*
	NAME....: MenuDemoRemoveEntry
	USAGE...: for EPLAN P8 (v2.09)
	AUTHOR..: S.Benner / BeDaSys
	VERSION.: 2011-05-11
	FUNC....: Demonstriert das dynamische Hinzufügen and Entfernen von Menüeinträgen per Script in EPlan P8
*/

/*
NAME....: MenuDemoRemoveEntry
USAGE ...: for EPLAN P8 (v2.09)
AUTHOR..: S.Benner / BeDaSys
VERSION.: 2011-05-11
FUNC....: Demonstrates the dynamic addition and removal of menu items via script in EPlan P8
*/

//
using Eplan.EplApi.Scripting;

public class AddMenu
{
    // Declarations 
    // -------------------------------------------------
    public static uint hndHMenu = new uint(); // Variable for the ID of the main menu
    public static uint hndMenuEntryL = new uint(); // Variable for the ID of the 2nd entry
    public static uint hndMenuEntryR = new uint(); // Variable for the ID of the 3rd entry
    public Eplan.EplApi.Gui.Menu DemoMainMenu = new Eplan.EplApi.Gui.Menu(); // The menu object
    //
    // Create the actions for the menu items
    // ------------------------------------------------ -
    //
    // Action: Switch to LEFT
    [DeclareAction("actLeft")] 
	public void actLeft() 
	{
        // Output message
        System.Windows.Forms.MessageBox.Show("Switch to LEFT");
        // Remove menu item "Left"
        DemoMainMenu.RemoveMenuItem(hndMenuEntryL);
        // Set menu entryID to 0

        hndMenuEntryL = 0;
        // Add menu item "Right" if it does not exist

        if (hndMenuEntryR == 0) {
			hndMenuEntryR = DemoMainMenu.AddMenuItem( // .AddMenuItem(
                "Right", // Entry name
                "actRight", // Item action,
                "Hereby I switch to right",//  Statustext,
				hndHMenu, //  Menu-ID,
				1, // Entry position(1 = rear or 0 = front),
				false, // Before seperator,
				false); // After seperator;
        }
	}
    //
    // Switch to RIGHT
    [DeclareAction("actRight")] 
	public void actRight()	
	{
		// Output message
		System.Windows.Forms.MessageBox.Show("Switch to RIGHT");
        // Remove menu item "Right"
        DemoMainMenu.RemoveMenuItem(hndMenuEntryR);
        // Set menu entryID to 0

        hndMenuEntryR = 0;
        // Add menu item "Left" if it does not exist

        if (hndMenuEntryL == 0) {
			hndMenuEntryL = DemoMainMenu.AddMenuItem(	// .AddMenuItem(
				"Left", // Entry name,
				"actLeft", // Item action,
                "Hereby I switch to Left",//  Statustext,
				hndHMenu, //  Menu-ID,
				1, // Entry position(1 = rear or 0 = front),
                false, // Before seperator,
                false); // After seperator;
        }
	}
    //
    // Switch to Left & Right
        [DeclareAction("actLeftRight")]
	public void actLeftRight()	
	{
		// Output message
		System.Windows.Forms.MessageBox.Show("Switch to Left & Right");
        // Add menu item "Left" if it does not exist
        if (hndMenuEntryL == 0) {
			hndMenuEntryL = DemoMainMenu.AddMenuItem(	// .AddMenuItem(
				"Left", //Entry name,
				"actLeft", // Item action,
				"Hereby I switch to Left",//  Statustext,
				hndHMenu, //  Menü-ID,
				1, // Entry position(1 = rear or 0 = front),
				false, // Before seperator,
				false); // After seperator);
			}
		// Menüeintrag "Right" hinzufügen falls er nicht vorhanden ist
		if (hndMenuEntryR == 0) {
			hndMenuEntryR = DemoMainMenu.AddMenuItem(	// .AddMenuItem(
				"Right", //Entry name,
				"actRight", // Item action,
				"Hereby I switch to Right",//  Statustext,
				hndHMenu, //  Menü-ID,
				1, // Entry position(1 = rear or 0 = front),
				false, // Before seperator,
				false); // After seperator);
			}
	}
	//	
	// Anlegen des Menüs
	// -------------------------------------------------
	[DeclareMenu]
	public void MenuFunction()
	{
        // Main menu includes Entry "Left and Right"

        hndHMenu = DemoMainMenu.AddMainMenu(    // .AddMainMenu(
            "Demo L / R switching",	// MenuName,
			"Window", //  RightNebenMenuName, 
			"Left and Right", // Entry name,
			"actLeftRight", // Item action,
            "Switch to Left & Right", // Statustext,
			1); //Entry position(1 = rear or 0 = front)		
		hndMenuEntryL = DemoMainMenu.AddMenuItem(	// .AddMenuItem(
			"Left", //Entry name,
			"actLeft", // Item action,
			"Hereby I switch to Left",//  Statustext,
			hndHMenu, //  Menü-ID,
			1, // Entry position(1 = rear or 0 = front),
			false, // Before seperator,
			false); // After seperator);
		hndMenuEntryR = DemoMainMenu.AddMenuItem(	// .AddMenuItem(
			"Right", //Entry name,
			"actRight", // Item action,
			"Hereby I switch to Right",//  Statustext,
			hndHMenu, //  Menü-ID,
			1, // Entry position(1 = rear or 0 = front),
			false, // Before seperator,
			false); // After seperator);
	}
}
