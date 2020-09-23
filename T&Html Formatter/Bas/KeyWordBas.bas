Attribute VB_Name = "KeyWordBas"
 
Option Explicit

Public KeyWords As CollC

Public Sub LoadKeyWords()
  Dim i As Integer
  
 Set KeyWords = New CollC
    
    With KeyWords
    
         .Add ".Circle", ".Local", ".Print", "Access", "AddressOf", "Alias", "And", "As", "As Date", "Base", "BF", "Binary", "Boolean", "ByRef", "Byte", "ByVal", _
               "Call", "Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "Circle", "Close", "Compare", "Const", "CSgn", "CSng", "CStr", _
               "Currency", "CVar", "CVErr", "Decimal", "Declare", "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng", "DefObj", "DefSng", _
               "DefStr", "DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "Empty", "End", "Enum", "EOF", "Eqv", "Erase", "Error", "Event", "Exit", "Explicit", _
               "False", "For", "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "In", "Input", "Integer", "Is", "LBound", "Let", "Lib", "Like", "Line", _
               "Local", "Lock", "Long", "Loop", "LSet", "Lof", "Mod", "Name", "New", "Next", "Not", "Nothing", "Object", "Open", "Option", "On", "Optional", "Or", "Output", _
               "Preserve", "Print", "Private", "Property", "Public", "Put", "Random", "RaiseEvent", "Read", "ReDim", "Resume", "Return", "RSet", "Seek", "Select", "Set", "Single", _
               "Spc", "Static", "Step", "String", "Stop", "Sub", "Tab", "Then", "To", "True", "Type", "TypeOf", "UBound", "Until", "Variant", "While", "Wend", "With", "Write", "Xor"
         .Sort
    End With
    
    BuildIndex
    
End Sub
