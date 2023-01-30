Function RepSymb(ByVal line_ As String, Optional dic_ As Variant, Optional check_doubles As Boolean) As String
    Dim IsDict As Boolean, CheckDoubles As Boolean, ResultLine As String, DecSec As String * 1
    'Checking parameters
    IsDict = Not IsMissing(dic_) And (StrComp(TypeName(dic_), "String") = 0)
    CheckDoubles = IIf(Not IsMissing(check_doubles), check_doubles, False) Or (IIf(Not IsMissing(dic_) And StrComp(TypeName(dic_), "String") <> 0, CBool(dic_), False))
    
    'Getting decimal separator from system
    DecSec = Application.DecimalSeparator
    
    'Creating RexExp object
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True: objRegExp.IgnoreCase = True
    
    'Checing if there's dictonary got from user and implement it in function
    If (IsDict) Then
        objRegExp.Pattern = "[" & dic_ & "]*"
        ResultLine = objRegExp.Replace(line_, "")
    Else
    
    'Using standart symbols ranges if there no dictonary got from user
        objRegExp.Pattern = "[^a-zA-Z0-9à-ÿÀ-ß" & DecSec & "]*"
        ResultLine = objRegExp.Replace(line_, "")
    End If
    
    'Checking if we need to remove repeated symbols and doing it
    If (CheckDoubles) Then
        objRegExp.Pattern = "(.)(?=\1)"
        ResultLine = objRegExp.Replace(ResultLine, "")
    End If
    
    'Returning clear string
    RepSymb = ResultLine
    
End Function