Attribute VB_Name = "Module1"
Sub CustomCitationMacro()
Attribute CustomCitationMacro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CustomCitationMacro Macro
' by Jeff Beuck, Cleveland State University
' v. 2.0 -- gets column positions by looking at field names so macro can
'           still run if fields are rearranged

Dim a As Long
Dim NumColumns As Integer
Dim NumRecords As Long
Dim CitationStr As String
Dim CitYear As String
Dim CitTitle As String
Dim CitJournal As String
Dim CitVolume As String
Dim CitIssue As String
Dim CitFirstPage As String
Dim CitLastPage As String
Dim CitDOI As String
Dim CitLastName(12) As String
Dim CitFirstInitial(12) As String
Dim CitMiddleInitial(12) As String
Dim ColOffset As Integer
Dim NextNameExists As Boolean
Dim TitleColumn As Integer
Dim YearColumn As Integer
Dim JournalColumn As Integer
Dim VolumeColumn As Integer
Dim IssueColumn As Integer
Dim FirstPageColumn As Integer
Dim LastPageColumn As Integer
Dim DOIColumn As Integer
Dim CitationColumn As Integer
Dim AuthorColumnsStart As Integer
Dim AuthorColumnsOffset As Integer
Dim AuthorFirstNameOffset As Integer
Dim AuthorMiddleOffset As Integer
Dim AuthorLastNameOffset As Integer
Dim Author1FirstNameColumn As Integer
Dim Author1MiddleNameColumn As Integer
Dim Author1LastNameColumn As Integer
Dim Author2FirstNameColumn As Integer
Dim Author2MiddleNameColumn As Integer
Dim Author2LastNameColumn As Integer
Dim NumAuthors As Integer
Dim NumCheck As Boolean
Dim LastAuthorReached As Boolean
Dim ErrorMessage As String

'ORIGINAL FORMULA FOR REFERENCE
'=(AB2)&", "&LEFT(Z2, 1)&"."&IF(AA2<>""," "&LEFT(AA2,1)&".","")&IF(AI2<>"",", "&IF(AP2="","& ","")&(AI2)&", "&LEFT(AG2, 1)&"."&IF(AH2<>""," "&LEFT(AH2,1)&".",""),"")&IF(AP2<>"",", "&IF(AW2="","& ","")&(AP2)&", "&LEFT(AN2, 1)&"."&IF(AO2<>""," "&LEFT(AO2,1)&".",""),"")&IF(AW2<>"",", "&IF(BD2="","& ","")&(AW2)&", "&LEFT(AU2, 1)&"."&IF(AV2<>""," "&LEFT(AV2,1)&".",""),"")&IF(BD2<>"",", "&IF(BK2="","& ","")&(BD2)&", "&LEFT(BB2, 1)&"."&IF(BC2<>""," "&LEFT(BC2,1)&".",""),"")&IF(BK2<>"",", "&IF(BR2="","& ","")&(BK2)&", "&LEFT(BI2, 1)&"."&IF(BJ2<>""," "&LEFT(BJ2,1)&".",""),"")&IF(BR2<>"",", "&IF(BY2="","& ","")&(BR2)&", "&LEFT(BP2, 1)&"."&IF(BQ2<>""," "&LEFT(BQ2,1)&".",""),"")&IF(BY2<>"",", "&IF(CF2="","& ","")&(BY2)&", "&LEFT(BW2, 1)&"."&IF(BX2<>""," "&LEFT(BX2,1)&".",""),"")&IF(CF2<>"",", "&IF(CM2="","& ","")&(CF2)&", "&LEFT(CD2, 1)&"."&IF(CE2<>""," "&LEFT(CE2,1)&".",""),"")&IF(CM2<>"",", "&IF(CT2="","& ","")&(CM2)&", "&LEFT(CK2, 1)&"."&IF(CL2<>""," "&LEFT(CL2,1)&".",""),"")&IF(CT2<>"",", "&IF(DA2="","& ","")&(CT2)
'&", "&LEFT(CR2, 1)&"."&IF(CS2<>""," "&LEFT(CS2,1)&".",""),"")&IF(DA2<>"",", &"&(DA2)&", "&LEFT(CY2, 1)&"."&IF(CZ2<>""," "&LEFT(CZ2,1)&".",""),"")&" ("&YEAR(C2)&"). "&A2&". "&E2&IF(O2<>"",", "&O2,"")&IF(P2<>"",", "&P2,"")&", "&Q2&"-"&R2&"."&IF(S2<>""," doi: "&S2,"")

'Get number of records and columns used in spreadsheet
NumRecords = ActiveSheet.UsedRange.Rows.Count
NumColumns = ActiveSheet.UsedRange.Columns.Count

'Initialize column variables
AuthorColumnsStart = 9999
AuthorColumnsOffset = 0
TitleColumn = 0
YearColumn = 0
JournalColumn = 0
VolumeColumn = 0
IssueColumn = 0
FirstPageColumn = 0
LastPageColumn = 0
DOIColumn = 0
CitationColumn = 0
Author1FirstNameColumn = 0
Author1MiddleNameColumn = 0
Author1LastNameColumn = 0
Author2FirstNameColumn = 0
Author2MiddleNameColumn = 0
Author2LastNameColumn = 0

'***************************************************************
'DETERMINE WHICH COLUMN MATCHES EACH FIELD NAME
'
'NOTE: if you get error box when running macro, replace
'fields names below with correct field names used in spreadsheet
'***************************************************************
For a = NumColumns To 1 Step -1
    fieldname = Trim(CStr(Cells(1, a).Value))
    If fieldname = "title" Then TitleColumn = a
    If fieldname = "publication_date" Then YearColumn = a
    If fieldname = "source_publication" Then JournalColumn = a
    If fieldname = "volnum" Then VolumeColumn = a
    If fieldname = "issnum" Then IssueColumn = a
    If fieldname = "fpage" Then FirstPageColumn = a
    If fieldname = "lpage" Then LastPageColumn = a
    If fieldname = "doi" Then DOIColumn = a
    If fieldname = "custom_citation" Then CitationColumn = a
    If InStr(fieldname, "author1") > 0 And InStr(fieldname, "fname") > 0 Then
        Author1FirstNameColumn = a
        If Author1FirstNameColumn < AuthorColumnsStart Then AuthorColumnsStart = Author1FirstNameColumn
    End If
    If InStr(fieldname, "author1") > 0 And InStr(fieldname, "mname") > 0 Then
        Author1MiddleNameColumn = a
        If Author1MiddleNameColumn < AuthorColumnsStart Then AuthorColumnsStart = Author1MiddleNameColumn
    End If
    If InStr(fieldname, "author1") > 0 And InStr(fieldname, "lname") > 0 Then
        Author1LastNameColumn = a
        If Author1LastNameColumn < AuthorColumnsStart Then AuthorColumnsStart = Author1LastNameColumn
    End If
    If InStr(fieldname, "author2") > 0 And InStr(fieldname, "fname") > 0 Then Author2FirstNameColumn = a
    If InStr(fieldname, "author2") > 0 And InStr(fieldname, "mname") > 0 Then Author2MiddleNameColumn = a
    If InStr(fieldname, "author2") > 0 And InStr(fieldname, "lname") > 0 Then Author2LastNameColumn = a
Next

'Display error alert if one or more fields couldn't be found
ErrorMessage = ""
If TitleColumn = 0 Then ErrorMessage = ErrorMessage & "Title column could not be found." & vbCr
If YearColumn = 0 Then ErrorMessage = ErrorMessage & "Year column could not be found." & vbCr
If JournalColumn = 0 Then ErrorMessage = ErrorMessage & "Journal column could not be found." & vbCr
If VolumeColumn = 0 Then ErrorMessage = ErrorMessage & "Volume column could not be found." & vbCr
If IssueColumn = 0 Then ErrorMessage = ErrorMessage & "Issue column could not be found." & vbCr
If FirstPageColumn = 0 Then ErrorMessage = ErrorMessage & "First Page column could not be found." & vbCr
If LastPageColumn = 0 Then ErrorMessage = ErrorMessage & "Last Name column could not be found." & vbCr
If DOIColumn = 0 Then ErrorMessage = ErrorMessage & "DOI column could not be found." & vbCr
If CitationColumn = 0 Then ErrorMessage = ErrorMessage & "Citation column could not be found." & vbCr
If Author1FirstNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 1 First Name column could not be found." & vbCr
If Author1MiddleNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 1 Middle Name column could not be found." & vbCr
If Author1LastNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 1 Last Name column could not be found." & vbCr
If Author2FirstNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 2 First Name column could not be found." & vbCr
If Author2MiddleNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 2 Middle Name column could not be found." & vbCr
If Author2LastNameColumn = 0 Then ErrorMessage = ErrorMessage & "Author 2 Last Name column could not be found." & vbCr

'Calculate the column offset between each set of author names
AuthorColumnsOffset = Author2FirstNameColumn - Author1FirstNameColumn
'Calculate the position of author first, middle, and last names in relation to each other
AuthorFirstNameOffset = Author1FirstNameColumn - AuthorColumnsStart
AuthorMiddleNameOffset = Author1MiddleNameColumn - AuthorColumnsStart
AuthorLastNameOffset = Author1LastNameColumn - AuthorColumnsStart

'Display error alert if author fields are not repeated in orderly fashion
If Author2LastNameColumn - Author1LastNameColumn <> AuthorColumnsOffset Or Author2MiddleNameColumn - Author1MiddleNameColumn <> AuthorColumnsOffset Then
    ErrorMessage = "Could not establish pattern for determining author fields." & vbCr
End If
If ErrorMessage <> "" Then
    ErrorMessage = "Error encountered." & vbCr & ErrorMessage
    MsgBox (ErrorMessage)
    End
End If


'Calculate number of author names
NumAuthors = 0
LastAuthorReached = False
Do Until LastAuthorReached = True
    NumCheck = False
    For a = 1 To NumColumns
        fieldname = Trim(CStr(Cells(1, a).Value))
        If InStr(fieldname, "author" & CStr(NumAuthors + 1)) > 0 Then NumCheck = True
    Next
    If NumCheck = True Then
        NumAuthors = NumAuthors + 1
    Else
        LastAuthorReached = True
    End If
Loop

'Cycle through each record on spreadsheet
For a = 2 To NumRecords
    CitationStr = ""
    
    'Get value from cell for each component of citation
    '"Cells(row, column).Value" retrieves value of cell
    CitYear = " (" & CStr(Year(Cells(a, YearColumn).Value)) & "). "
    CitTitle = CStr(Cells(a, TitleColumn).Value) & ". "
    CitJournal = CStr(Cells(a, JournalColumn).Value)
    CitVolume = CStr(Cells(a, VolumeColumn).Value)
        If CitVolume <> "" Then CitVolume = ", " & CitVolume
    CitIssue = CStr(Cells(a, IssueColumn).Value)
        If CitIssue <> "" Then CitIssue = "(" & CitIssue & ")"
    CitFirstPage = CStr(Cells(a, FirstPageColumn).Value)
    CitLastPage = CStr(Cells(a, LastPageColumn).Value)
    CitDOI = CStr(Cells(a, DOIColumn).Value)
        If CitDOI <> "" Then CitDOI = " doi: " & CitDOI
    CitLastName(1) = CStr(Cells(a, Author1LastNameColumn).Value)
    CitFirstInitial(1) = Left(CStr(Cells(a, Author1FirstNameColumn).Value), 1) & "."
    CitMiddleInitial(1) = Left(CStr(Cells(a, Author1MiddleNameColumn).Value), 1)
       If CitMiddleInitial(1) <> "" Then CitMiddleInitial(1) = " " & CitMiddleInitial(1) & "."
  
    'Start building citation string (add name #1)
    CitationStr = CitLastName(1) & ", " & CitFirstInitial(1) & CitMiddleInitial(1)
    'Add names #2-12 by looping through possible name fields
    For b = 2 To NumAuthors
        ColOffset = (b - 1) * AuthorColumnsOffset
        'Get name information for name #b
            'OLD CODE: CitLastName(b) = CStr(Cells(a, (28 + ColOffset)).Value)
        CitLastName(b) = CStr(Cells(a, (AuthorColumnsStart + AuthorLastNameOffset + ColOffset)).Value)
            'OLD CODE: CitFirstInitial(b) = Left(CStr(Cells(a, (26 + ColOffset)).Value), 1) & "."
        CitFirstInitial(b) = Left(CStr(Cells(a, (AuthorColumnsStart + AuthorFirstNameOffset + ColOffset)).Value), 1) & "."
            'OLD CODE: CitMiddleInitial(b) = Left(CStr(Cells(a, (27 + ColOffset)).Value), 1)
        CitMiddleInitial(b) = Left(CStr(Cells(a, (AuthorColumnsStart + AuthorMiddleNameOffset + ColOffset)).Value), 1)
            If CitMiddleInitial(b) <> "" Then CitMiddleInitial(b) = " " & CitMiddleInitial(b) & "."
            
        'Check to see if there is any value in the cell for the next name
        If CStr(Cells(a, (AuthorColumnsStart + AuthorLastNameOffset + ColOffset + AuthorColumnsOffset))) <> "" Then NextNameExists = True Else NextNameExists = False
        
        'Add name #b to the citation string
        If CitLastName(b) <> "" Then
            CitationStr = CitationStr & ", "
            'If there is no name after this, add an ampersand before it
            If NextNameExists = False Then CitationStr = CitationStr & "& "
            'Actually add name #b to the string
            CitationStr = CitationStr & CitLastName(b) & ", " & CitFirstInitial(b) & CitMiddleInitial(b)
        End If
       
    Next
    'Finish off citation string
    CitationStr = CitationStr & CitYear & CitTitle & CitJournal & CitVolume & CitIssue & ", " & CitFirstPage & "-" & CitLastPage & "." & CitDOI
    
    'Set citation string as value of cell (11 = Row K)
    Cells(a, CitationColumn).Value = CitationStr
Next

End Sub
