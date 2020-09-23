VB-Win32DLL Compiler, V1.01 build 041212 (Rev. 1)
Made by KFGG, China.P.R, 12/12/2004


Instructions:

  This small project allows you to create a EXE or DLL through VB 
that could export functions and subs, and you can use your own 
file just like using WinApi. (e.g. Declare MyFunction Lib "MyDLL.dll" Alias "MyFunction"...)
  See the sample project "\Arithmetic".


Usage:

  Rename VB's original Link.EXE to LinkMS.EXE, then build Link.exe and put 
it into the same directory as the original Link.EXE.
  When you want to compile a exportable DLL or EXE, just press SHIFT, 
and do not release until the interface shows when building in VB.
  Specify the module or other files that contain functions and sub to export, 
and specify the File to exporting declares, click "Compile Win32 DLL", 
you will have your own VB-Compiled exportable DLL or EXE. You can also 
modify the parameters and then compile as parameter.


Tips:

  1.You can create a EXE project just like the sample, and build it as a DLL.
Also, you can use ActiveX DLL project, it can build as a exportable DLL too.
  2.I suggest that every variable in the defines of the functions or subs accessing 
through "ByVal", it would be much safer to your project.


Author's information:

  Mail me (kfggstudio@hotmail.com) if you have any suggestions or 
problems, please note that i cannot promise that the project has 
no bugs or that is completely safe. If you develop the project, 
please send me one copy.Many thanks.


Version's history:

  12/05/2004 Version 1.00
    The first version.
  12/12/2004 Version 1.01 (Rev. 1)
    This version corrects a bug that do not support multiple line defines, and there're
    small changes with the function MakeParameter and other parts, modifies some of the
    spellings in the project.
