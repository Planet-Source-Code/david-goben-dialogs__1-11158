<div align="center">

## Dialogs


</div>

### Description

This is one of over a hundred modules I have developed for getting my work done faster. This module display various system dialog boxes to configure COM ports, printer ports, get the default printer, view printer properties, and view document properties.
 
### More Info
 
This module should be saved as a .BAS file. I called it modPrinter.bas, but that was when it was a piece of baby code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Goben](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-goben.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-goben-dialogs__1-11158/archive/master.zip)

### API Declarations

```
'*************************************************
' API calls, constants, and types
'*************************************************
' size of a device name string
Private Const CCHDEVICENAME = 32
' size of a form name string
Private Const CCHFORMNAME = 32
Private Const DM_IN_PROMPT = 4
Private Const DM_OUT_BUFFER = 2
Public Type DEVMODE
 dmDeviceName As String * CCHDEVICENAME 'name of the printer
 dmSpecVersion As Integer        'DEVMODE version
 dmDriverVersion As Integer       'printer driver version
 dmSize As Integer            'total size of DEVMODE w/o private data
 dmDriverExtra As Integer        'total size of private data
 dmFields As Long            'flags indicating which fields are valid
 dmOrientation As Integer        'portraint/landscape (see DMORIENT_xxx)
 dmPaperSize As Integer         'papersize (see DMPAPER_xxx)
 dmPaperLength As Integer        'paper length in tenths of mm's
 dmPaperWidth As Integer         'paper width in tenths of mm's
 dmScale As Integer           'scales paper size by x/100
 dmCopies As Integer           'number of copies
 dmDefaultSource As Integer       'reserved. keep at zero
 dmPrintQuality As Integer        'qualiyt (see DMRES_xxx) (or horz res DPI)
 dmColor As Integer           'color type (see DMCOLOR_xx)
 dmDuplex As Integer           'reserved
 dmYResolution As Integer        'if not 0, vert res in DPI
 dmTTOption As Integer          'How to print TT fonts (see DTT_xxx)
 dmCollate As Integer          'collation (see DMCOLLATE_xxx)
 dmFormName As String * CCHFORMNAME   'NT only. Name of printer form to use
 dmUnusedPadding As Integer       'reserved
 dmBitsPerPel As Integer         'bits per pixel for display (not printers)
 dmPelsWidth As Long           'width of display in pixels (not printers)
 dmPelsHeight As Long          'height of display in pixels (not printers)
 dmDisplayFlags As Long         'DM_GRAYSCALE or SM_INTERLACED *not printers)
 dmDisplayFrequency As Long       'Display frequency (not printers)
 dmICMMethod As Long           'one of the DMICM_xxx constants (color matching)
 dmICMIntent As Long           'one of the DMICM_xxx constants (intensity)
 dmMediaType As Long           'one of the DMMEDIA_xxx constants
 dmDitherType As Long          'on of the DMDITHER_xxx constants
 dmReserved1 As Long           'reserved
 dmReserved2 As Long           'reserved
End Type
Private Type PRINTER_DEFAULTS
 pDatatype As String
 pDevMode As Long
 DesiredAccess As Long
End Type
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, ByVal pDevModeInput As Long) As Long
Private Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hWnd As Long, ByVal Flags As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pPortName As String) As Long
'customized calls
Private Declare Function DocumentPropertiesStr Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As String, ByVal pDevModeInput As String, ByVal fMode As Long) As Long
Private Declare Sub CopyMemoryDM Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As DEVMODE, ByVal source As String, ByVal Length As Long)
```


### Source Code

```
'*************************************************
' modPrinterDialogs:
' This module displays a number of dialogs, which
' are provided by the following functions:
'
' ConfigureCOMPort():   Configure the specified COM port number (1-4)
' ConfigureLPTPort():   Configure the specified Printer port number (1-4)
' ConfigureAPort():    Configure a specified port
' GetDefaultPrinter():   This function retrieves the definition
'             of the default printer on this system
' ViewPrinterProperties(): View/change printer properties dialog
' ViewDocProperties():   View/change document properties
' ConnectToAPrinter():   Connect to a local/network printer
'
'EXAMPLES:
' Dim dm As DEVMODE         'used to gather data by ViewDocProperties()
'
' Call ConfigureAPort(Me, "COM2:") 'configure COM port 2
' Call ConfigureCOMPort(Me, 2)   'configure COM port 2
' Call ConfigureLPTPort(Me, 1)   'configure LPT port 1
' Debug.Print GetDefaultPrinter   'display default printer name, device, port
' Call ViewPrinterProperties(Me)  'view/change default printer's properties
' Call ConnectToAPrinter(Me)    'connect to a local/network printer
' Call ViewDocProperties(Me, dm)  'set up document printing options.
' Debug.Print "Copies = " & dm.dmCopies
' Debug.Print "Orientation = " & dm.dmOrientation
' Debug.Print "Quality = " & dm.dmPrintQuality
'*************************************************
''''INSERT API/Global goodies here
'*************************************************
' ConfigureCOMPort(): Configure the specified COM port number (1-4)
'*************************************************
Public Function ConfigureCOMPort(Frm As Form, PortNumber As Integer)
 ConfigureCOMPort = ConfigurePort("", Frm.hWnd, "COM" & CStr(PortNumber) & ":")
End Function
'*************************************************
' ConfigureLPTPort(): Configure the specified Printer port number (1-4)
'*************************************************
Public Function ConfigureLPTPort(Frm As Form, PortNumber As Integer)
 ConfigureLPTPort = ConfigurePort("", Frm.hWnd, "LPT" & CStr(PortNumber) & ":")
End Function
'*************************************************
' ConfigureAPort(): Configure a specified port
'*************************************************
Public Function ConfigureAPort(Frm As Form, PortName As String)
 ConfigureAPort = ConfigurePort("", Frm.hWnd, UCase$(PortName))
End Function
'*************************************************
' ViewPrinterProperties(): View/change printer properties dialog
'*************************************************
Public Sub ViewPrinterProperties(Frm As Form, Optional PrtDevice As String = "")
  Dim hPrinter As Long
  hPrinter& = OpenAPrinter(PrtDevice)
  If hPrinter = 0 Then
    If PrtDevice = "" Then
     MsgBox "Unable to open default printer"
    Else
     MsgBox "Unable to open " & PrtDevice & " printer"
    End If
    Exit Sub
  End If
  Call PrinterProperties(Frm.hWnd, hPrinter)
  Call ClosePrinter(hPrinter)
End Sub
'*************************************************
' ViewDocProperties(): View/change document properties
'*************************************************
Public Sub ViewDocProperties(Frm As Form, MyDevMode As DEVMODE, Optional DeviceName As String = "")
  Dim bufsize As Long, res As Long
  Dim dmInBuf As String
  Dim dmOutBuf As String
  Dim hPrinter As Long
  hPrinter = OpenAPrinter(DeviceName)
  If hPrinter = 0 Then
   If DeviceName = "" Then
    MsgBox "Unable to open default printer"
   Else
    MsgBox "Unable to open " & DeviceName & " printer"
   End If
   Exit Sub
  End If
  ' The output DEVMODE structure will reflect any changes
  ' made by the printer setup dialog box.
  ' Note that no changes will be made to the default
  ' printer settings!
  bufsize = DocumentProperties(Frm.hWnd, hPrinter, DeviceName, 0, 0, 0)
  dmInBuf = String(bufsize, 0)
  dmOutBuf = String(bufsize, 0)
  res = DocumentPropertiesStr(Frm.hWnd, hPrinter, DeviceName, dmOutBuf, dmInBuf, DM_IN_PROMPT Or DM_OUT_BUFFER)
  ' Copy the data buffer into the DEVMODE structure
  CopyMemoryDM MyDevMode, dmOutBuf, Len(MyDevMode)
ClosePrinter hPrinter
End Sub
'*************************************************
' ConnectToAPrinter(): Connect to a local/network printer
'*************************************************
Public Sub ConnectToAPrinter(Frm As Form)
 Call ConnectToPrinterDlg(Frm.hWnd, 0)
End Sub
'*************************************************
' GetDefaultPrinter(): This function retrieves the definition
'           of the default printer on this system
'*************************************************
Public Function GetDefaultPrinter() As String
  Dim def As String
  Dim di As Long
  def = String(128, 0)
  di = GetProfileString("WINDOWS", "DEVICE", "", def, 127)
  If di Then GetDefaultPrinter = Left$(def, di - 1)
End Function
'*************************************************
' OpenAPrinter(): open a printer (default or user-specified)
'*************************************************
Private Function OpenAPrinter(Optional DeviceName As String = "") As Long
  Dim dev$, devname As String, devoutput As String
  Dim hPrinter As Long, res As Long
  Dim pdefs As PRINTER_DEFAULTS
  pdefs.pDatatype = vbNullString
  pdefs.pDevMode = 0
  pdefs.DesiredAccess = PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE
  If DeviceName = "" Then
   dev = GetDefaultPrinter() ' Get default printer info
   If dev = "" Then Exit Function
   DeviceName = GetDeviceName(dev)
  End If
  devname = DeviceName
  ' You can use OpenPrinterBynum to pass a zero as the
  ' third parameter, but you won't have full access to
  ' edit the printer properties
  res = OpenPrinter(devname, hPrinter, pdefs)
  If res <> 0 Then OpenAPrinter = hPrinter
End Function
'*************************************************
'  Retrieves the name portion of a device string
'*************************************************
Private Function GetDeviceName(dev As String) As String
  Dim npos As Integer
  npos = InStr(dev, ",")
  GetDeviceName = Left$(dev, npos - 1)
End Function
```

