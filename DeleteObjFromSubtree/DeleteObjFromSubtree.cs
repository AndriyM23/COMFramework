/* 
@echo off && cls

set WinDirNet=%WinDir%\Microsoft.NET\Framework

IF EXIST "%WinDirNet%\v2.0.50727\csc.exe"   set CSharpCompiler=%WinDirNet%\v2.0.50727\csc.exe
IF EXIST "%WinDirNet%\v3.5\csc.exe"         set CSharpCompiler=%WinDirNet%\v3.5\csc.exe
IF EXIST "%WinDirNet%\v4.0.30319\csc.exe"   set CSharpCompiler=%WinDirNet%\v4.0.30319\csc.exe

"%CSharpCompiler%" /nologo /define:DONT_HOLD_PROGRAM_ON_END /out:"%~0.exe" %0

"%~0.exe" %1 %2 %3 %4

del "%~0.exe" 

pause
exit 
*/ 
  
using System;
using System.Diagnostics;
using System.IO;

public enum TypesOfObjectsToDelete {All, OnlyFiles, OnlyDirectories};

class MainClass 
{ 
    static int Main(string[] args)
    {
        //System.Console.WriteLine("C# compiler version: " + System.Environment.Version);
        
        if (args.Length == 0)
        {
            System.Console.WriteLine("Usage: " + Process.GetCurrentProcess().ProcessName + " <dirOrFileNameOrWldcrdToDelete> [/a|/f|/d] [/s] [<SubtreeTopName>]"); 
            System.Console.WriteLine("If <SubtreeTopName> not specified - the location of this batch:");
            System.Console.WriteLine(System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]));
            System.Console.WriteLine("will be taken as top.");
            System.Console.WriteLine("[/a|/f|/d] - /a - delete dirs&files(default), /f - delete files, /f - delete dirs");
            System.Console.WriteLine("[/s] - simulate (only print what will be deleted, but not delete).");
            #if (!DONT_HOLD_PROGRAM_ON_END)
            System.Console.ReadLine();
            #endif
            return 0;
        }
        
        string objectNameToDelete="";
        string ourScriptDirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
        string subtreeRootToLookIn = ourScriptDirectory;
        bool bSimulate = false;
        TypesOfObjectsToDelete whatToDelete = TypesOfObjectsToDelete.All;
        int pathParamsCounter = 0;
        foreach (string arg in args)
        {
            if (arg[0] != '/' && arg[0] != '-')
            {
                ++pathParamsCounter;
                if (pathParamsCounter == 1)
                    objectNameToDelete = arg;
                else if (pathParamsCounter == 2)
                {
                    string combinedPath = Path.Combine(ourScriptDirectory, arg);
                    subtreeRootToLookIn = Path.GetFullPath(combinedPath);
                }
            }
            else if (arg.ToLower() == "/s" || arg.ToLower() == "-s")
                bSimulate = true;
            else if (arg.ToLower() == "/d" || arg.ToLower() == "-d")
                whatToDelete = TypesOfObjectsToDelete.OnlyDirectories;
            else if (arg.ToLower() == "/f" || arg.ToLower() == "-f")
                whatToDelete = TypesOfObjectsToDelete.OnlyFiles;
        }
        
        if (objectNameToDelete.Length == 0)
        {
            System.Console.WriteLine("Usage: " + Process.GetCurrentProcess().ProcessName + " <dirOrFileNameOrWldcrdToDelete> [/a|/f|/d] [/s] [<SubtreeTopName>]"); 
            System.Console.WriteLine("If <SubtreeTopName> not specified - the location of this batch:");
            System.Console.WriteLine(System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]));
            System.Console.WriteLine("will be taken as top.");
            System.Console.WriteLine("[/a|/f|/d] - /a - delete dirs&files(default), /f - delete files, /f - delete dirs");
            System.Console.WriteLine("[/s] - simulate (only print what will be deleted, but not delete).");
            #if (!DONT_HOLD_PROGRAM_ON_END)
            System.Console.ReadLine();
            #endif
            return 0;
        }
            
        System.Console.WriteLine("{0} all {1} named \"{2}\" from subtree rooted in:\n{3}", 
                                 bSimulate ? "Showing" : "Removing",
                                 whatToDelete==TypesOfObjectsToDelete.All ? "objects" : (whatToDelete==TypesOfObjectsToDelete.OnlyFiles ? "files" : "directories"),
                                 objectNameToDelete,
                                 subtreeRootToLookIn );
        
        RemoveObjectsFromSubtree(objectNameToDelete, subtreeRootToLookIn, whatToDelete, bSimulate);

        #if (!DONT_HOLD_PROGRAM_ON_END)
        System.Console.ReadLine();
        #endif
        return 0;
    }
    
    static void RemoveObjectsFromSubtree(string objectNameToDelete, string subtreeRootToLookIn, TypesOfObjectsToDelete whatToDelete, bool bSimulate = true)
    {
        if (whatToDelete == TypesOfObjectsToDelete.All || whatToDelete == TypesOfObjectsToDelete.OnlyFiles)
        {
            string [] filesToDelete = System.IO.Directory.GetFiles(subtreeRootToLookIn, objectNameToDelete, System.IO.SearchOption.AllDirectories);
            Console.WriteLine("\n{0} files:\n{1}", bSimulate ? "Showing" : "Deleting", string.Join("\n", filesToDelete));
            
            if (!bSimulate)
                foreach(string fileName in filesToDelete)
                    File.Delete(fileName);
        }
        
        if (whatToDelete == TypesOfObjectsToDelete.All || whatToDelete == TypesOfObjectsToDelete.OnlyDirectories)
        {
            string [] dirsToDelete = System.IO.Directory.GetDirectories(subtreeRootToLookIn, objectNameToDelete, System.IO.SearchOption.AllDirectories);
            Console.WriteLine("\n{0} directories:\n{1}", bSimulate ? "Showing" : "Deleting", string.Join("\n", dirsToDelete));
            
            if (!bSimulate)
                foreach(string dirName in dirsToDelete)
                    Directory.Delete(dirName, true);            
        }
    }
} 
