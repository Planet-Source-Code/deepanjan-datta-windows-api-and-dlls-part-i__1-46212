<div align="center">

## Windows API and DLLs \-\-\- Part\-I


</div>

### Description

An introduction to Windows API and DLLs!!! Part II on PSC!!! Part III coming soon
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Deepanjan Datta](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/deepanjan-datta.md)
**Level**          |Intermediate
**User Rating**    |3.1 (25 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/deepanjan-datta-windows-api-and-dlls-part-i__1-46212/archive/master.zip)





### Source Code

<html>
<p align="center"><font color="#FF0000" size="6"><span style="background-color: #FFFF00"><b><i><u>Windows
API</u></i></b></span></font></p>
<p align="center"> </p>
<p align="left"><b><i>Full form </i></b>: API --- Application Programming
Interface</p>
<p align="left">                  
DLL --- Dynamic Link Library</p>
<p align="center"> </p>
<p align="left">The Windows API is a collection of routines available to you,
the Visual Basic programmer. In a way, these API routines are like internal
functions of Visual Basic.</p>
<p align="left"> So many Windows API routines exist that just about
anything you can do from Windows, you can do from a Visual Basic application by
calling the appropriate Windows API routine.</p>
<p align="left">All Windows API routines are stored in files called DLLs.
Several thousand API routines are available for use.</p>
<p align="left"> </p>
<p align="left"><b><i>Note </i></b>: Most DLL files have '.DLL' extension.</p>
<p align="left">Any program you write has access to the Windows DLLs.</p>
<p align="left"> </p>
<p align="left"><b><i>Following are the three most common DLLs </i>:</b></p>
<ul>
 <li>
 <p align="left"><b>USER32.DLL</b> --- Contains functions that control the
 Windows environment and the user's interface, such as cursors, menus,
 windows etc.</li>
 <li>
 <p align="left"><b>GDI32.DLL </b>--- Contains functions that control output
 to the screen and other devices.</li>
 <li>
 <p align="left"><b>KERNEL32.DLL </b>--- Contains functions that control the
 internal Windows hardware and software interface.</li>
</ul>
<p align="left">There are other DLLs such as COMDLG.DLL, MAPI32.DLL,
NETAPI32.DLL, WINMM.DLL etc.</p>
<p align="left"> </p>
<p align="center"><font color="#0000FF" size="4"><u><i><b>Using the 'Declare'
statement</b></i></u></font></p>
<p align="left">Calling Windows API routines requires a statement called
'Declare'.</p>
<p align="left">The 'Declare' statement performs the following tasks :</p>
<ul>
 <li>
 <p align="left">Specifies where the API function is located</li>
 <li>
 <p align="left">Identifies arguments needed by the API function by number
 and data type</li>
 <li>
 <p align="left">Specifies whether or not the API function returns a value</li>
</ul>
<p align="left">The following format describes the subroutine procedure version
of the 'Declare' statement :</p>
<table border="1" width="100%">
 <tr>
 <td width="100%"><font face="Arial">Declare Sub procName Lib
 "libName" [Alias "alias"] [([ByVal] var1 [As dataType]
 [, [ByVal] var2 [As dataType]] ... [, [ByVal] varN [As dataType])]</font></td>
 </tr>
</table>
<p align="left"> Here are two examples :</p>
<table border="1" width="100%">
 <tr>
 <td width="100%"><font face="Arial">Declare Function GetWindowsDirectory Lib
 "kernel32" Alias "GetWindowsDirectoryA"_</font>
 <p><font face="Arial">(ByVal lpBuffer As String, ByVal nSize As Long) As
 Long</font></td>
 </tr>
</table>
<table border="1" width="100%">
 <tr>
 <td width="100%"><font face="Arial">Declare Sub GetSystemInfo Lib
 "kernel32" (lpSystemInfo As SystemInfo)</font></td>
 </tr>
</table>
<p>Here is an example for calling a simple API routine:</p>
<p>This example sounds the speaker</p>
<table border="1" width="104%">
 <tr>
 <td width="100%">Private Declare Function MessageBeep Lib "user32"
 (ByVal wType As Long) As Long
 <p>Private Sub cmdBeep_Click() 'You need to have a command button named
 cmdBeep for this example to work</p>
 <p>Dim Beeper As Variant</p>
 <p>Beeper=MessageBeep(1)</p>
 <p>End Sub</td>
 </tr>
</table>
</html>

