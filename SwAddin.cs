using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;


namespace airDucts
{
	/// <summary>
	/// Summary description for airDucts.
	/// </summary>
	[Guid("bf243245-6ade-46e4-8871-8620875d5073"), ComVisible(true)]
	[SwAddin(
		Description = "airDucts description",
		Title = "airDucts",
		LoadAtStartup = true
		)]
	public class SwAddin : ISwAddin
	{
		#region Local Variables
		ISldWorks iSwApp = null;
		ICommandManager iCmdMgr = null;
		int addinID = 0;
		BitmapHandler iBmp;

		public const int mainCmdGroupID = 1;
		public const int mainItemID1 = 0;
		

		#region Event Handler Variables
		Hashtable openDocs = new Hashtable();
		SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
		#endregion

		


		// Public Properties
		public ISldWorks SwApp
		{
			get { return iSwApp; }
		}
		public ICommandManager CmdMgr
		{
			get { return iCmdMgr; }
		}

		public Hashtable OpenDocs
		{
			get { return openDocs; }
		}

		#endregion

		#region SolidWorks Registration
		[ComRegisterFunctionAttribute]
		public static void RegisterFunction(Type t)
		{
			#region Get Custom Attribute: SwAddinAttribute
			SwAddinAttribute SWattr = null;
			Type type = typeof(SwAddin);

			foreach (System.Attribute attr in type.GetCustomAttributes(false))
			{
				if (attr is SwAddinAttribute)
				{
					SWattr = attr as SwAddinAttribute;
					break;
				}
			}

			#endregion

			try
			{
				Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
				Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

				string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
				Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
				addinkey.SetValue(null, 0);

				addinkey.SetValue("Description", SWattr.Description);
				addinkey.SetValue("Title", SWattr.Title);

				keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
				addinkey = hkcu.CreateSubKey(keyname);
				addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
			}
			catch (System.NullReferenceException nl)
			{
				Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
				System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
			}

			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);

				System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
			}
		}

		[ComUnregisterFunctionAttribute]
		public static void UnregisterFunction(Type t)
		{
			try
			{
				Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
				Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

				string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
				hklm.DeleteSubKey(keyname);

				keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
				hkcu.DeleteSubKey(keyname);
			}
			catch (System.NullReferenceException nl)
			{
				Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
				System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
			}
			catch (System.Exception e)
			{
				Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
				System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
			}
		}

		#endregion

		#region ISwAddin Implementation
		public SwAddin()
		{
		}

		public bool ConnectToSW(object ThisSW, int cookie)
		{
			iSwApp = (ISldWorks)ThisSW;
			addinID = cookie;

			//Setup callbacks
			iSwApp.SetAddinCallbackInfo(0, this, addinID);

			#region Setup the Command Manager
			iCmdMgr = iSwApp.GetCommandManager(cookie);
			AddCommandMgr();
			#endregion

			#region Setup the Event Handlers
			SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
			openDocs = new Hashtable();
			AttachEventHandlers();
			#endregion

			return true;
		}

		public bool DisconnectFromSW()
		{
			RemoveCommandMgr();
			DetachEventHandlers();

			System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
			iCmdMgr = null;
			System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
			iSwApp = null;
			//The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
			GC.Collect();
			GC.WaitForPendingFinalizers();

			GC.Collect();
			GC.WaitForPendingFinalizers();

			return true;
		}
		#endregion

		#region UI Methods
		public void AddCommandMgr()
		{
			ICommandGroup cmdGroup;
			if (iBmp == null)
				iBmp = new BitmapHandler();
			Assembly thisAssembly;
			int cmdIndex0;
			string Title = "Воздуховоды", ToolTip = "Воздуховоды и фасонные элементы";


			int[] docTypes = new int[] { (int)swDocumentTypes_e.swDocASSEMBLY };
									   //(int)swDocumentTypes_e.swDocDRAWING,
									   //(int)swDocumentTypes_e.swDocPART};

			thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


			int cmdGroupErr = 0;
			bool ignorePrevious = false;

			object registryIDs;
			//get the ID information stored in the registry
			bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

			int[] knownIDs = new int[1] { mainItemID1};

			if (getDataResult)
			{
				if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
				{
					ignorePrevious = true;
				}
			}

			cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
			cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("airDucts.ToolbarLarge.bmp", thisAssembly);
			cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("airDucts.ToolbarSmall.bmp", thisAssembly);
			cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("airDucts.MainIconLarge.bmp", thisAssembly);
			cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("airDucts.MainIconSmall.bmp", thisAssembly);

			int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
			cmdIndex0 = cmdGroup.AddCommandItem2("AirDucts", -1, "Создание воздуховодов и фасонных элементов ", 
				"Воздуховоды и фасонные элементы", 0, "AirDucts", "", mainItemID1, menuToolbarOption);

			cmdGroup.HasToolbar = true;
			cmdGroup.HasMenu = true;
			cmdGroup.Activate();

			bool bResult;

			foreach (int type in docTypes)
			{
				CommandTab cmdTab;

				cmdTab = iCmdMgr.GetCommandTab(type, Title);

				if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
				{
					bool res = iCmdMgr.RemoveCommandTab(cmdTab);
					cmdTab = null;
				}

				//if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
				if (cmdTab == null)
				{
					cmdTab = iCmdMgr.AddCommandTab(type, Title);

					CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

					int[] cmdIDs = new int[1];
					int[] TextType = new int[1];

					cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

					TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

					bResult = cmdBox.AddCommands(cmdIDs, TextType);



					CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
					cmdIDs = new int[1];
					TextType = new int[1];

					
					bResult = cmdBox1.AddCommands(cmdIDs, TextType);

					cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);

				}

			}
			thisAssembly = null;

		}

		public void RemoveCommandMgr()
		{
			iBmp.Dispose();

			iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
			
		}

		public bool CompareIDs(int[] storedIDs, int[] addinIDs)
		{
			List<int> storedList = new List<int>(storedIDs);
			List<int> addinList = new List<int>(addinIDs);

			addinList.Sort();
			storedList.Sort();

			if (addinList.Count != storedList.Count)
			{
				return false;
			}
			else
			{

				for (int i = 0; i < addinList.Count; i++)
				{
					if (addinList[i] != storedList[i])
					{
						return false;
					}
				}
			}
			return true;
		}



		#endregion

		#region UI Callbacks

		public void AirDucts()
		{
			string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart); //мб здесь
			if ((partTemplate != null) && (partTemplate != ""))
			{
				//IModelDoc2 modDoc = iSwApp.IActiveDoc2;
				//startForm start = new startForm();
				Form1 form = new Form1();
				form.ShowDialog();
				form.Dispose();

			}
			else
			{
				System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options " +
					"and make sure there is a part template selected, or select a new part template.");
			}

		}


	
		#endregion

		#region Event Methods
		public bool AttachEventHandlers()
		{
			AttachSwEvents();
			//Listen for events on all currently open docs
			AttachEventsToAllDocuments();
			return true;
		}

		private bool AttachSwEvents()
		{
			try
			{
				SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
				SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
				SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
				SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
				SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
				return true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				return false;
			}
		}



		private bool DetachSwEvents()
		{
			try
			{
				SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
				SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
				SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
				SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
				SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
				return true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				return false;
			}

		}

		public void AttachEventsToAllDocuments()
		{
			ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
			while (modDoc != null)
			{
				if (!openDocs.Contains(modDoc))
				{
					AttachModelDocEventHandler(modDoc);
				}
				else if (openDocs.Contains(modDoc))
				{
					bool connected = false;
					DocumentEventHandler docHandler = (DocumentEventHandler)openDocs[modDoc];
					if (docHandler != null)
					{
						connected = docHandler.ConnectModelViews();
					}
				}
				modDoc = (ModelDoc2)modDoc.GetNext();
			}
		}

		public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
		{
			if (modDoc == null)
				return false;

			DocumentEventHandler docHandler = null;

			if (!openDocs.Contains(modDoc))
			{
				switch (modDoc.GetType())
				{
					case (int)swDocumentTypes_e.swDocPART:
						{
							docHandler = new PartEventHandler(modDoc, this);
							break;
						}
					case (int)swDocumentTypes_e.swDocASSEMBLY:
						{
							docHandler = new AssemblyEventHandler(modDoc, this);
							break;
						}
					case (int)swDocumentTypes_e.swDocDRAWING:
						{
							docHandler = new DrawingEventHandler(modDoc, this);
							break;
						}
					default:
						{
							return false; //Unsupported document type
						}
				}
				docHandler.AttachEventHandlers();
				openDocs.Add(modDoc, docHandler);
			}
			return true;
		}

		public bool DetachModelEventHandler(ModelDoc2 modDoc)
		{
			DocumentEventHandler docHandler;
			docHandler = (DocumentEventHandler)openDocs[modDoc];
			openDocs.Remove(modDoc);
			modDoc = null;
			docHandler = null;
			return true;
		}

		public bool DetachEventHandlers()
		{
			DetachSwEvents();

			//Close events on all currently open docs
			DocumentEventHandler docHandler;
			int numKeys = openDocs.Count;
			object[] keys = new Object[numKeys];

			//Remove all document event handlers
			openDocs.Keys.CopyTo(keys, 0);
			foreach (ModelDoc2 key in keys)
			{
				docHandler = (DocumentEventHandler)openDocs[key];
				docHandler.DetachEventHandlers(); //This also removes the pair from the hash
				docHandler = null;
			}
			return true;
		}
		#endregion

		#region Event Handlers
		//Events
		public int OnDocChange()
		{
			return 0;
		}

		public int OnDocLoad(string docTitle, string docPath)
		{
			return 0;
		}

		int FileOpenPostNotify(string FileName)
		{
			AttachEventsToAllDocuments();
			return 0;
		}

		public int OnFileNew(object newDoc, int docType, string templateName)
		{
			AttachEventsToAllDocuments();
			return 0;
		}

		public int OnModelChange()
		{
			return 0;
		}

		#endregion
	}

}
