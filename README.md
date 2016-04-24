# dynamic-fre-wrapper
C# .NET 4.5 library that provides a dynamic IDisposable wrapper for ABBYY FineReader Engine 11 COM object and all its sub-objects. COM object types are determined at runtime so you do not need to reference the corresponding .NET Interop library in your project and recompile it when new maintenance release of ABBYY FineReader Engine 11 is available. Just replace old Engine distribution files, (re-)register FREngine.tlb together with FREngine.dll using the following command and your code works again:

**regsvr32 /n /i:"path to FREngine.tlb *folder*" "path to FREngine.dll *file*"**

If you want to unregister these files, you can simply add /u option to the command:

**regsvr32 */u* /n /i:"path to FREngine.tlb *folder*" "path to FREngine.dll *file*"**

The repository contains Visual Studio 2013 solution with a sample project illustrating the library usage.

In order to run the sample, you must have a valid ABBYY FineReader Engine 11 license and know your ProjectID which is passed to DynamicFrEngine() constructor.
If you have a local license file, you must also specify the correct password in the constructor and place the license into "%ProgramData%\ABBYY\SDK\11\Licenses" folder.

Since all wrapped Engine COM objects support IDisposable interface, you may want to use this capability to have more predictable control over releasing COM objects. This approach may be helpful if you observe that the way .NET garbage collector works is not suitable for your specific workflow. When you place the COM object into “using” block or explicitly call its Dispose() method, Marshal.FinalReleaseComObject() is called internally for this COM object.

Engine collection types support IEnumerable.

NOTE. When using this dynamic wrapper, IntelliSense technology is not available. Therefore, please apply to ABBYY FineReader Engine Help files regarding method names and signatures in order to call them correctly from your code.
