VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
'       CLASS VARIABLES
''''''''''''''''''''''''''''''''''''''''''
'   Variables from Userform
''''''''''''''''''''''''''''''''''''''''''
Private pName As String
Private pDate As Date
Private pDepartment As String
Private pCategory1 As String
Private pCategory2 As String
Private pCategorty3 As String
Private pPriority As String
Private pSpCat1 As String
Private pSpCat2 As String
Private pSpCat3 As String
Private pGenDescr As String

''''''''''''''''''''''''''''''''''''''''''
'   Variable Locations for
'   Database and Form input
''''''''''''''''''''''''''''''''''''''''''
Private pCellLocations(1, 10) As Integer
Private Enum eCellLoc
    DBName = 1
    DBDate = 2
    DBDepartment = 3
    DBCategory1 = 4
    DBCategory2 = 5
    DBCategory3 = 6
    DBPriority = 7
    DBSpCat1 = 8
    DBSpCat2 = 9
    DBSpCat3 = 10
    DbGenDescr = 11
End Enum



''''''''''''''''''''''''''''''''''''''''''
'   NAME PROPERTY
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(Value As String)
    pName = Value
End Property

''''''''''''''''''''''''''''''''''''''''''
'   DATE PROPERTY
''''''''''''''''''''''''''''''''''''''''''
'Public Property Get Date() As Date
'    Date = pDate
'End Property
'Public Property Let Date(Value As String)
'    pDate = Value
'End Property

''''''''''''''''''''''''''''''''''''''''''
'   DEPARTMENT PROPERTY
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Department() As String
    Department = pDepartment
End Property
Public Property Let Department(Value As String)
    pDepartment = Value
End Property

''''''''''''''''''''''''''''''''''''''''''
'   Y/N CATEGORY PROPERTIES
''''''''''''''''''''''''''''''''''''''''''
'Category1
Public Property Get Category1() As String
    Category1 = pCategory1
End Property
Public Property Let Category1(Value As String)
    pCategory1 = Value
End Property
'Category2
Public Property Get Category2() As String
    Category2 = pCategory2
End Property
Public Property Let Category2(Value As String)
    pCategory2 = Value
End Property
'Category3
Public Property Get Category3() As String
    Category3 = pCategory3
End Property
Public Property Let Category3(Value As String)
    pCategory3 = Value
End Property

''''''''''''''''''''''''''''''''''''''''''
'   PRIORITY PROPERTY
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Priority() As String
    Priority = pPriority
End Property
Public Property Let Priority(Value As String)
    pPriority = Value
End Property

''''''''''''''''''''''''''''''''''''''''''
'   SPECIFIC CATEGORY PROPERTIES
''''''''''''''''''''''''''''''''''''''''''
'Category1
Public Property Get SpCat1() As String
    pSpCat1 = pSpCat1
End Property
Public Property Let pSpCat1(Value As String)
    pSpCat1 = Value
End Property
'Category2
Public Property Get pSpCat2() As String
    pSpCat2 = pSpCat2
End Property
Public Property Let pSpCat1(Value As String)
    pSpCat2 = Value
End Property
'Category3
Public Property Get pSpCat3() As String
    pSpCat3 = pSpCat3
End Property
Public Property Let pSpCat3(Value As String)
    pSpCat3 = Value
End Property

''''''''''''''''''''''''''''''''''''''''''
'   GENERAL DESCRIPTION PROPERTY
''''''''''''''''''''''''''''''''''''''''''
Public Property Get GenDescr() As String
    GenDescr = pGenDescr
End Property
Public Property Let GenDescr(Value As String)
    pGenDescr = Value
End Property

'================================================
'       CLASS FUNCTIONS

