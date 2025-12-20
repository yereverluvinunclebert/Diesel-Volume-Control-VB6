Attribute VB_Name = "mSupport"
Option Explicit

Public Declare Function PSGetPropertyDescription Lib "propsys.dll" (PropKey As PROPERTYKEY, riid As UUID, ppv As Any) As Long
'Public Declare Function PSFormatForDisplay Lib "propsys.dll" (propkey As PROPERTYKEY, propvar As Variant, pdfFlags As PROPDESC_FORMAT_FLAGS, pwszText As Long, ByVal cchText As Long) As Long
Public Declare Function PSFormatPropertyValue Lib "propsys.dll" (ByVal pps As Long, ByVal ppd As Long, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As Long) As Long
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long) ' Frees memory allocated by the shell
Public Declare Function PropVariantToVariant Lib "propsys" (ByRef propvar As Any, ByRef var As Variant) As Long
Public Declare Function PSGetNameFromPropertyKey Lib "propsys.dll" (PropKey As PROPERTYKEY, ppszCanonicalName As Long) As Long

Public Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function


Public Sub dbg_enumstore(pPStore As IPropertyStore)
Dim isif As IShellItem2
Dim pidlt As Long
Dim pProp As IPropertyDescription
Dim pk As PROPERTYKEY
Dim lpe As Long
Dim lpProp As Long
Dim i As Long, j As Long
Dim vProp As Variant
Dim vrProp As Variant
Dim vte As VbVarType
Dim sPrName As String
Dim sFmtProp As String
'Call CoInitialize(0)
'Create a reference to IShellItem2
If (pPStore Is Nothing) Then
    Debug.Print "Failed to get IPropertyStore"
    Exit Sub
End If

'Get the number of properties
pPStore.GetCount lpe
Debug.Print "Total number of properties=" & lpe

On Error GoTo eper
For i = 0 To (lpe - 1)
    'Loop through each property; starting with information about which property we're working with
'    Debug.Print "trace pos 1"
    pPStore.GetAt i, pk
'    Debug.Print "trace pos 2"
    PSGetNameFromPropertyKey pk, lpProp
    Debug.Print "trace pos 3"
    sPrName = LPWSTRtoStr(lpProp)
    Debug.Print "Property Name=" & sPrName & ",SCID={" & Hex$(pk.fmtid.Data1) & "-" & Hex$(pk.fmtid.Data2) & "-" & Hex$(pk.fmtid.Data3) & "-" & Hex$(pk.fmtid.Data4(0)) & Hex$(pk.fmtid.Data4(1)) & "-" & Hex$(pk.fmtid.Data4(2)) & Hex$(pk.fmtid.Data4(3)) & Hex$(pk.fmtid.Data4(4)) & Hex$(pk.fmtid.Data4(5)) & Hex$(pk.fmtid.Data4(6)) & Hex$(pk.fmtid.Data4(7)) & "}, " & pk.pid
    
    'Some properties don't return a name; if you don't catch that it leads to a full appcrash
    If Len(sPrName) > 1 Then
        'PSFormatPropertyValue takes the raw data and formats it according to the current locale
        'Using these APIs lets us completely avoid dealing with PROPVARIANT, a huge bonus.
        'If you don't need the raw data, this is all it takes
'        Debug.Print "trace pos 4"
       PSGetPropertyDescription pk, IID_IPropertyDescription, pProp
        If (pProp Is Nothing) = False Then
'        Debug.Print "trace pos 4a"
        PSFormatPropertyValue ObjPtr(pPStore), ObjPtr(pProp), PDFF_DEFAULT, lpProp
        Else
            Debug.Print "pprop=nothing, no psformat.."
        End If
        sFmtProp = LPWSTRtoStr(lpProp)
        Debug.Print "Formatted value=" & sFmtProp
    Else
        Debug.Print "Unknown Propkey; can't get formatted value"
    End If
'     Debug.Print "trace pos 5"
   If pk.pid = 14 Then
    Debug.Print "dispval i=" & i
    'Now we'll display the raw data
    pPStore.GetValue pk, vProp
    Debug.Print "trace pos 6"
    PropVariantToVariant vProp, vrProp 'PROPVARIANT is exceptionally difficult to work with in VB, but at
                                       'least for file properties this seems to work for most

    vte = VarType(vrProp)
    If (vte And vbArray) = vbArray Then 'this always seems to be vbString and vbArray, haven't encountered other types
        For j = LBound(vrProp) To UBound(vrProp)
            Debug.Print "Value(" & j & ")=" & CStr(vrProp(j))
        Next j
    Else
    Select Case vte
        Case vbDataObject, vbObject, vbUserDefinedType
            Debug.Print "<cannot display this type>"
        Case vbEmpty, vbNull
            Debug.Print "<empty or null>"
        Case vbError
            Debug.Print "<vbError>"
        Case Else
            Debug.Print "Value=" & CStr(vrProp)
    End Select
    End If
    lpProp = 0
    Set pProp = Nothing
    VariantClear vProp
    End If
Next i
Exit Sub
eper:
    Debug.Print "Property conversion error->" & Err.Description
    Resume Next


End Sub
