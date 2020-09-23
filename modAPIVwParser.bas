Attribute VB_Name = "modAPIVwParser"
Option Explicit
' Compatible with v2 & v3 files for Win16, Win32 & WinCE versions

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Enum apiSections ' used when querying for details on an apv file
    [_minsec] = 0
    apDeclarations = 0
    apTypes = 1
    apConstants = 2
    apEnumerations = 3
    [_maxsec] = 3
End Enum
Public Enum SearchTypeStructure ' used for searching apvs
    [_minsrch] = 0
    apExactMatch = 0
    apBeginsWith = 1
    apiContains = 2
    [_maxsrch] = 2
End Enum
Private Type apiQuickIndexArray ' indexes
    apiIndexArray() As Long
End Type
    
Private sourceFile As String            ' apv file name & path
Private filenum As Integer              ' keep apv file open for faster access

' we will index evey 500th apv record
' Smaller groupings mean quicker searches, but more memory cached
Private Const quickIndexGroup = 500

' The total number of sub/functions, constants, types, enumerations, etc
Private apiTotals(0 To 3) As Long
' Indexes to the 500th occurrences of apv records (0 to 3 are for the sections)
Private apiQuickIndex(0 To 3) As apiQuickIndexArray
' Alternate indexes to indicate where alphabetical entries begin
' 0 to 3 for sections, 1-26 for A-Z, 0 for ASCII<A & 27 for ASCII>z
Private apiQuickAlpha(0 To 3, 0 To 27) As Long

Public Property Let APIfile(FilePathName As String)
    
    Dim fnr As Integer, I As Integer
    
    If FilePathName = "" Then
        ' clean up
        If filenum Then
            Close #filenum
            filenum = 0
            sourceFile = ""
            Erase apiTotals
            Erase apiQuickAlpha
            For I = 0 To UBound(apiQuickAlpha, 1)
                Erase apiQuickIndex(I).apiIndexArray
            Next
        End If
    
    Else
        ' test new apv file before replacing existing one if applicable
        On Error GoTo NoFile
        ' see if error occurs when opening
        fnr = FreeFile()
        Open FilePathName For Input Access Read As fnr
        Close #fnr
        
        ' no error, open in binary mode
        fnr = FreeFile()
        Open FilePathName For Binary Access Read As #fnr
        
        ' process indexes & offsets
        If InitializeOffsets(fnr) Then
            sourceFile = FilePathName
            filenum = fnr
        Else
            ' error occurred, wrong file, failed access or something
            On Error GoTo 0
            Err.Raise 53, "SetAPIviewer Filename", "Invalid APIviewer Filename"
        End If

    End If
    
Exit Property

NoFile:
Err.Raise Err.Number, "SetAPIviewer Filename", Err.Description
End Property

Public Property Get APIfile() As String
    APIfile = sourceFile
End Property

Public Property Get APIsectionCount(Section As apiSections) As Long
    If Section < [_minsec] Or Section > [_maxsec] Then Exit Property
    
    ' return number of items in the requested apv section
    APIsectionCount = apiTotals(Section)

End Property

Private Function InitializeOffsets(fileNr As Integer) As Boolean

' version 3+ byte definitions
' 5  << &H1 major version
' 8  << &H4 number of subs (1346)
' 12 << &H4 number of functions (4834)
' 16 << &H4 number of constants (52933)
' 20 << &H4 number of types (469)

' v2
' 24 << 1st delcaration text length
' 26 << begin 1st declaration text
' 26+lenText is next declaration length

' v3
' 24 << &H4 number of enumerations (4)
' 40 << 1st delcaration text length
' 42 << begin 1st declaration text
' 42+lenText is next declaration length

' ^^ each item (both versions) is always preceded by the length of the item (2 bytes)

If filenum Then Close #filenum

Dim byteOffset As Long, Looper As Long
Dim apiCount As Long, apiLen As Integer
Dim sectionIdx As Long
Dim apiBytes() As Byte
Dim fileVer As Byte
Dim qArrayPtr As Long, qDefPtr As Long

Const byteBase As Byte = 6 '5th byte really, but Get# is 1 bound vs 0 bound

On Error GoTo BadFile

ReDim apiBytes(0 To 23)
Get #fileNr, byteBase, apiBytes()
    
     ' check version
    CopyMemory fileVer, apiBytes(0), &H1
    If fileVer = 0 Or fileVer > 3 Then
        Close #fileNr
        Exit Function
    End If
    
    ' clear any previous arrays & then repopulate
    Erase apiQuickAlpha()
    Erase apiTotals
    
    ' apv has subs & function totals as separate entries
    CopyMemory apiTotals(apDeclarations), apiBytes(3), &H4
    CopyMemory apiTotals(apTypes), apiBytes(7), &H4
    ' combine the two
    apiTotals(apDeclarations) = apiTotals(apDeclarations) + apiTotals(apTypes)
    ' prep arrays
    Erase apiQuickIndex(apDeclarations).apiIndexArray
    ReDim apiQuickIndex(apDeclarations).apiIndexArray(0 To 2, 0 To apiTotals(apDeclarations) \ quickIndexGroup)
    
    ' get constant total & prep arrays
    CopyMemory apiTotals(apConstants), apiBytes(11), &H4
    Erase apiQuickIndex(apConstants).apiIndexArray
    ReDim apiQuickIndex(apConstants).apiIndexArray(0 To 2, 0 To apiTotals(apConstants) \ quickIndexGroup)
    
    ' get type total & prep arrays
    CopyMemory apiTotals(apTypes), apiBytes(15), &H4
    Erase apiQuickIndex(apTypes).apiIndexArray
    ReDim apiQuickIndex(apTypes).apiIndexArray(0 To 2, 0 To apiTotals(apTypes) \ quickIndexGroup)
    
    ' get enumeration total & prep arrays if applicable
    Erase apiQuickIndex(apEnumerations).apiIndexArray
    If fileVer > 2 Then
        CopyMemory apiTotals(apEnumerations), apiBytes(19), &H4
        ReDim apiQuickIndex(apEnumerations).apiIndexArray(0 To 2, 0 To apiTotals(apEnumerations) \ quickIndexGroup)
    Else
        ' ver 2 didn't have enumerations
        ReDim apiQuickIndex(apEnumerations).apiIndexArray(0 To 2, 0 To 0)
    End If

' now walk the rest of the file to find offsets
ReDim apiBytes(0 To 2)
' 0-1 used to read integer length values
' 2 used to reach ASCII value of section item

' set hardcoded offset where first declaration item is found
byteOffset = Choose(fileVer, 25, 25, 41)

' array pointers to 2 different arrays
qArrayPtr = -1
qDefPtr = apiTotals(apDeclarations)

For Looper = 1 To UBound(apiTotals) * 2 + 2
    For apiCount = 0 To apiTotals(sectionIdx) - 1
    
        If Looper Mod 2 Then
            ' cache location of every 500th record (section index)
            If apiCount \ quickIndexGroup > qArrayPtr Then
                qArrayPtr = qArrayPtr + 1
                apiQuickIndex(sectionIdx).apiIndexArray(0, qArrayPtr) = apiCount
                apiQuickIndex(sectionIdx).apiIndexArray(1, qArrayPtr) = byteOffset
            End If
        Else
            ' cache location of every 500th record (section data)
            If apiCount \ quickIndexGroup > qDefPtr Then
                qDefPtr = qDefPtr + 1
                apiQuickIndex(sectionIdx).apiIndexArray(2, qDefPtr) = byteOffset
            End If
        End If
        
        ' get the integer length & ascii value of 1st character
        Get #fileNr, byteOffset, apiBytes()
        CopyMemory apiLen, apiBytes(0), &H2
        
        If Looper Mod 2 Then
            ' here we will also cache index of the 1st entry of each Alphabetical item (A-Z)
            Select Case apiBytes(2)
            Case 65 To 90
                If apiQuickAlpha(sectionIdx, apiBytes(2) - 64) = 0 Then apiQuickAlpha(sectionIdx, apiBytes(2) - 64) = apiCount + 1
            Case 97 To 122
                If apiQuickAlpha(sectionIdx, apiBytes(2) - 96) = 0 Then apiQuickAlpha(sectionIdx, apiBytes(2) - 96) = apiCount + 1
            Case Is < 65
                If apiQuickAlpha(sectionIdx, 0) = 0 Then apiQuickAlpha(sectionIdx, 0) = apiCount + 1
            Case Else
                If apiQuickAlpha(sectionIdx, 27) = 0 Then apiQuickAlpha(sectionIdx, 27) = apiCount + 1
            End Select
        End If
        
        ' increment location of next item
        byteOffset = byteOffset + apiLen + 2
    Next
    
    ' set appropriate pointers & section refs
    Select Case Looper
    Case 1: ' declarations
        sectionIdx = apDeclarations
        qDefPtr = -1
    Case 2, 3: ' types
        sectionIdx = apTypes
        If Looper = 2 Then qArrayPtr = -1 Else qDefPtr = -1
    Case 4, 5: ' constants
        sectionIdx = apConstants
        If Looper = 4 Then qArrayPtr = -1 Else qDefPtr = -1
    Case 6, 7: ' enumerations
        sectionIdx = apEnumerations
        If Looper = 6 Then qArrayPtr = -1 Else qDefPtr = -1
    End Select

Next

' return success
InitializeOffsets = True
Exit Function

BadFile:
' return failure
Close #fileNr
End Function

Public Function GetApiListing(Section As apiSections, _
    Optional ByVal fromIndex As Long = 0, _
    Optional ByVal toIndex As Long = -1) As String()

' function called to return an array of apv index items

' Section :: Declarations, Types, Constants or Enumerations
' fromIndex :: 0 bound & optional.
'   This is the first item in the apv to return
' toIndex :: 0 bound & optional.
'   This is the last item in the apv to return. -1 returns all

' Example of returning the 1st 100 Declaration names
'    Dim I As Integer
'    Dim vNames() As String
'    vNames = GetApiListing(apDeclarations, 0, 99)
'    ' prove we did something
'    For I = 0 To UBound(vNames)
'        ' if less than 100 were available, those would be nullstrings
'        If Len(vanmes(I)) = 0 Then Exit For
'        Debug.Print vNames(I)
'    Next


Dim rtnArray() As String
ReDim rtnArray(-1 To -1)

If filenum = 0 Then Exit Function
If fromIndex > apiTotals(Section) - 1 Then Exit Function

Dim byteOffset As Long, apiOffset As Long
Dim apiCount As Long
Dim apiLen As Integer
Dim apiN(0 To 1) As Byte
Dim apiBytes() As Byte

' determine where to start based off of fromIndex
apiCount = fromIndex \ quickIndexGroup
' get the byteOffset of the nth 500th item
byteOffset = apiQuickIndex(Section).apiIndexArray(1, apiCount)
' get the index of the nth 500th item
apiOffset = apiQuickIndex(Section).apiIndexArray(0, apiCount)

' fix the toIndex as needed
If toIndex < 0 Then
    toIndex = apiTotals(Section)
ElseIf toIndex > apiTotals(Section) Then
    toIndex = apiTotals(Section)
End If

If apiTotals(Section) Then
    ' move to the first index item
    MoveToIndex apiOffset, fromIndex, byteOffset
    
    ' prep the return array
    ReDim rtnArray(0 To toIndex - fromIndex)
    
    For apiCount = fromIndex To toIndex
        
        ' get the length of the item
        Get #filenum, byteOffset, apiN()
        CopyMemory apiLen, apiN(0), &H2
        ' prep array to receive that item & then get it
        ReDim apiBytes(0 To apiLen - 1)
        Get #filenum, , apiBytes
        
        ' cache the item
        rtnArray(apiCount - fromIndex) = StrConv(apiBytes, vbUnicode)
        
        ' track byte offsets
        byteOffset = byteOffset + apiLen + 2
    
    Next

End If
GetApiListing = rtnArray
End Function


Public Function ParseDBsection(Section As apiSections, Index As Long, apiName As String) As String

' parse a selected section item

' Section :: Declaration, Constant, Type or Enumeration
' Index :: the 0 bound Index of the item to parse
' apiName :: the caption of the item to parse (passing it prevents re-reading)
'   if a null string is passed, the parameter will contain the name
'   for the Index passed when the function returns

If filenum = 0 Then Exit Function
If Index > apiTotals(Section) - 1 Then Exit Function

Dim apiCount As Long
Dim byteOffset As Long
Dim apiData As String
Dim apiLen As Integer
Dim apiFormat As String
Dim I As Integer
Dim apiBytes() As Byte, apiParts() As String

' calcuate data byte offset of the 500th record containing this Index
apiCount = Index \ quickIndexGroup

' didn't pass a name, let's get it based off of the index
If Len(apiName) = 0 Then
    ' same procedure used below & remarked there, not here
    byteOffset = apiQuickIndex(Section).apiIndexArray(1, apiCount)
    MoveToIndex apiQuickIndex(Section).apiIndexArray(0, apiCount), Index, byteOffset
    
    ReDim apiBytes(0 To 1)
    Get #filenum, byteOffset, apiBytes
    CopyMemory apiLen, apiBytes(0), &H2
    
    ReDim apiBytes(0 To apiLen - 1)
    Get #filenum, , apiBytes
    apiName = StrConv(apiBytes, vbUnicode)

End If

' calcuate data byte offset of the 500th record containing this Index
byteOffset = apiQuickIndex(Section).apiIndexArray(2, apiCount)

' move to the proper byte offset
MoveToIndex apiQuickIndex(Section).apiIndexArray(0, apiCount), Index, byteOffset

' prep arrays for reading the data & get the data
ReDim apiBytes(0 To 1)
Get #filenum, byteOffset, apiBytes
CopyMemory apiLen, apiBytes(0), &H2

ReDim apiBytes(0 To apiLen - 1)
Get #filenum, , apiBytes
apiData = StrConv(apiBytes, vbUnicode)

Select Case Section

Case apDeclarations 'format declarations
    ' sample apv definition:  "gdi32" (?hdc&)&
    ' when done parsing it will look like:
    ' Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
    
    ' apv files have a flag at end of data to indicate if function or sub
    ' the flag is generally &, %, $ if a function; otherwise, no flag is present
    If Right$("*" & apiData, 1) = ")" Then
        apiData = "Declare Sub " & apiName & " Lib " & apiData
    Else
        apiData = "Declare Function " & apiName & " Lib " & apiData
    End If

    ' all entries have the dll name enclosed by quotes
    I = InStr(apiData, Chr$(34))
    If I Then
    
        ' version 3 added the .dll part if it didn't exist in version 2
        If InStr(apiData, ".") = 0 Then ' add the .dll to statement
            I = InStr(I + 1, apiData, Chr$(34))
            apiData = Left$(apiData, I - 1) & ".dll" & Mid$(apiData, I)
        End If
        
        ' now find the beginning of the parameters
        I = InStr(I + 1, apiData, "(")
        ' split parameters based on commas
        apiParts = Split(Mid$(apiData, I + 1), ",")
        ' remove the split portion & we will add it back later
        apiData = Left$(apiData, I)
        
        For I = 0 To UBound(apiParts)
            Select Case Left$(LTrim(apiParts(I)) & ")", 1)
            Case ")"    ' no parameters, usually a Sub vs Function
            Case "?"    ' ByVal reference
                apiParts(I) = Replace$(LTrim$(apiParts(I)), "?", "ByVal ")
            Case "~"    ' ByRef reference
                apiParts(I) = Replace$(LTrim$(apiParts(I)), "~", "ByRef ")
            Case Else
                ' v3 added ByRef if ByVal not specifically annotated
                apiParts(I) = "ByRef " & Mid$(LTrim$(apiParts(I)), 1)
            End Select
        Next
        
        ' now put back the parts
        apiData = apiData & Join(apiParts, ", ")
        ' these symbols need to be replaced within the coded declaration
        apiData = Replace$(apiData, "&", " As Long")
        apiData = Replace$(apiData, "$", " As String")
        apiData = Replace$(apiData, "%", " As Integer")
        'apiData = Replace$(apiData, "#", " As Double") '<< not currently used
        'apiData = Replace$(apiData, "!", " As Single") '<< not currently used
        'apiData = Replace$(apiData, "@", " As Currency") '<< not currently used
        
    Else
        ' bad data in the database
        apiData = ""
    End If
    
Case apTypes
    ' types are easy. The already have the tabs & vbCrLfs encoded in the apv
    apiData = "Type " & apiName & vbCrLf & apiData & vbCrLf & "End Type"
   
Case apEnumerations
    ' enumerations are identical to Types
    apiData = "Enum " & apiName & vbCrLf & apiData & vbCrLf & "End Enum"
    
Case apConstants
    ' routine the apiviewer uses is flawed...
    ' It reports the wrong Type: Long vs String for...
    '   Const wszNAMESEPARATORDEFAULT As **Long** = szNAMESEPARATORDEFAULT
    '   where szNAMESEPARATORDEFAULT is declared As **String** = "\n"
    ' The right way would be to look up the referenced constant to be absolutely sure
    
    If InStr(apiData, Chr$(34)) Then
        ' since constants can be long & string mixed, if the definition
        ' has a quote in it, it will be a String constant
        apiData = "Const " & apiName & " As String = " & apiData
    Else
        ' here we want to test so we can overcome the error in the API Viewer
        apiParts = Split(apiData, " ")
        For I = 0 To UBound(apiParts)
            ' TODO:
        Next
        apiData = "Const " & apiName & " As Long = " & apiData
    End If
End Select
ParseDBsection = apiData
End Function


Public Function GetAlphaIndex(Section As apiSections, ByVal KeyCode As Byte) As Long

' function returns the Index of the 1st apv Item that begins with passed ASCII value
' Primarily used for scrolling thru a listing. This does not return any byte offsets

' Section :: Declaration, Type, Constant or Enumerations
' KeyCode :: ASCII value to return first occurrence

Dim lRtn As Long
lRtn = -1

If KeyCode = 0 Then Exit Function
If apiTotals(Section) Then
    
    ' ensure upper case is used here
    KeyCode = Asc(UCase(Chr$(KeyCode)))
    
    ' identify which Index we want to look at (A-Z)
    Select Case KeyCode
    Case 65 To 90
        KeyCode = KeyCode - 64
    Case Is < 65
        KeyCode = 0
    Case Else
        KeyCode = 27
    End Select
    
    ' return the Index
    lRtn = apiQuickAlpha(Section, KeyCode) - 1

End If

GetAlphaIndex = lRtn
End Function

Public Function SearchAPIFile(Section As apiSections, Criteria As String, _
    SearchMode As SearchTypeStructure, Optional ByVal fromIndex As Long = -1, _
    Optional byteOffset As Long = 0) As Long

' Function returns index of match based on type of criteria & search mode.
' Return value when
'   SeachMode = apExactMatch :: the next Index that can be searched
'   SearchMode = apiContains :: the next Index that can be searched
'   SearchMode = apBeginsWith :: the same Index of the match
' The byteOffset parameter returned will be for the function's return value Index


' By using the return value and byteOffset return value,
' a fast recursive search can be performed

' Section :: Declaration, Type, Constant, Enumerations
' Criteria :: search string (not case sensitive)
' fromIndex :: optional. The Index to begin the search on
' byteOffset :: optional. If used, must be accruate or bad data will be returned
    ' when function returns and a match is found, byteOffset will be
    ' the correct location of that match which can then be sent back into
    ' this routine to find the next match, etc, etc, etc
    
' Example of Recursive Seach from Outside this module
' This example will return all Constants beginning with WM_

'    Dim I As Integer
'    Dim nextOffset As Long, nextIndex As Long
'    Dim wmMsgs() As String, wmTitle() As String
'
'    ReDim wmMsgs(-1 To -1)
'    ' do the first search with unknown starting index & unknown offset
'    nextIndex = SearchAPIFile(apConstants, "WM_", apiContains, -1, nextOffset)
'
'    Do While nextIndex > -1
'       ' resize arrays to accept new entry
'       ReDim Preserve wmMsgs(-1 To UBound(wmMsgs) + 1)
'       ReDim Preserve wmTitle(0 To UBound(wmMsgs) + 1)
'       ' notice we will pass a nullstring wmTitle(Ubound(wmTitle)) &
'       ' have the ParseDBsection function fill it in for us
'       wmMsgs(UBound(wmMsgs)) = ParseDBsection(apConstants, _
'            nextIndex - 1, wmTitle(UBound(wmTitle)))
'       ' continue the search using now known index & offsets
'       nextIndex = SearchAPIFile(apConstants, "WM_", apiContains, nextIndex, nextOffset)
'    Loop
'    ' prove we did something
'    For I = 0 To UBound(wmMsgs)
'        Debug.Print wmTitle(I); "  "; wmMsgs(I)
'    Next



Dim lRtn As Long

lRtn = -1
If filenum = 0 Then Exit Function
If fromIndex > apiTotals(Section) - 1 Then Exit Function
If Len(Criteria) = 0 Then Exit Function
If SearchMode < [_minsrch] Or SearchMode > [_maxsrch] Then Exit Function

If apiTotals(Section) Then

    Dim apiCount As Long
    Dim apiLen As Integer
    Dim apiName As String
    Dim apiN(0 To 1) As Byte
    Dim apiBytes() As Byte

    
    If byteOffset = 0 Then  ' unknown byte offset
        If fromIndex < 0 Then
            ' starting index not provided, get the offset for that Index
            If SearchMode = apiContains Then
                ' partial match search, start from the beginning
                byteOffset = apiQuickIndex(Section).apiIndexArray(1, 0)
                fromIndex = 0
            
            Else
                ' exact match or match that begins with crtieria, find first alpha Index
                fromIndex = GetAlphaIndex(Section, Asc(UCase(Left$(Criteria & Chr$(0), 1))))
                If fromIndex < 0 Then Exit Function
                
                ' the previous function returns an index, not byte offsets
                ' find the byte offset for the Index
                apiCount = fromIndex \ quickIndexGroup
                byteOffset = apiQuickIndex(Section).apiIndexArray(1, apiCount)
                MoveToIndex apiQuickIndex(Section).apiIndexArray(0, apiCount), fromIndex, byteOffset
            
            End If
        Else
            ' Index provided, but not a byte offset. Get that offset
            apiCount = fromIndex \ quickIndexGroup
            byteOffset = apiQuickIndex(Section).apiIndexArray(1, apiCount)
            MoveToIndex 0, fromIndex, byteOffset
        End If
    End If
    
    ' now find a match if possible
    For apiCount = fromIndex To apiTotals(Section) - 1
        
        ' get the integer length of the item, then get the item
        Get #filenum, byteOffset, apiN()
        CopyMemory apiLen, apiN(0), &H2
        ReDim apiBytes(0 To apiLen - 1)
        Get #filenum, , apiBytes
        
        ' depending on search mode, we compare
        Select Case SearchMode
        Case apiContains    ' partial search
            If InStr(1, StrConv(apiBytes, vbUnicode), Criteria, vbTextCompare) Then
                lRtn = apiCount + 1
                byteOffset = byteOffset + 2 + apiLen
                Exit For
            End If
        Case apExactMatch   ' exact match only
            Select Case StrComp(StrConv(apiBytes, vbUnicode), Criteria, vbTextCompare)
            Case 0 ' exact match
                lRtn = apiCount + 1
                byteOffset = byteOffset + 2 + apiLen
                Exit For
            Case Is > 0 ' passed up alphabetically without a match
                Exit For
            End Select
        Case apBeginsWith ' "begins with" match
            Select Case StrComp(Left$(StrConv(apiBytes, vbUnicode), Len(Criteria)), Criteria, vbTextCompare)
            Case 0 ' exact match
                lRtn = apiCount
                Exit For
            Case Is > 0 ' passed up alphabetically without a match
                Exit For
            End Select
        ' can add other search methods like "Ends With", etc
        End Select
        
        ' increment byte offsets & continue looping
        byteOffset = byteOffset + 2 + apiLen
    Next
End If

' return success or failure
SearchAPIFile = lRtn

End Function

Private Sub MoveToIndex(startIndex As Long, endIndex As Long, byteOffset As Long)

' Helper function which moves to a specific index from a near byte offset

Dim apiCount As Long
Dim apiLen As Integer
Dim apiN(0 To 1) As Byte

For apiCount = startIndex To endIndex - 1
    Get #filenum, byteOffset, apiN()
    CopyMemory apiLen, apiN(0), &H2
    byteOffset = byteOffset + apiLen + 2
Next

End Sub

