# dynamic-fre-wrapper
C# .NET 4.5 class library that provides a dynamic IDisposable wrapper around ABBYY FineReader Engine 11 COM object and all derived COM objects. All accessed COM objects should be placed into "using" block to avoid huge memory consumption when iterating through document data as .NET garbage collector in no hurry to clean up heavy COM objects like document pages. The repository contains Visual Studio 2013 solution with a sample project of the library usage.

In order to run the sample you should have a valid ABBYY FineReader Engine 11 license and know your ProjectID to pass into DynamicFrEngine() constructor. If you have a local license file, you should also specify a password in the constructor and place the license into "%ProgramData%\ABBYY\SDK\11\Licenses" folder.

You should register "FREngine.tlb" together with Inproc and Outproc FREngine COM Servers using the following command before running the sample:

regsvr32 /n /i:"path to FREngine.tlb **folder**" "path to FREngine.dll **file**"

After tests, you can roll things back simply adding /u option to the command:

regsvr32 /u /n /i:"path to FREngine.tlb **folder**" "path to FREngine.dll **file**"
