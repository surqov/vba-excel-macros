Function RepSymb(ByVal line_ As String, Optional dic_ As Variant, Optional check_doubles As Boolean) As String
    Dim IsDict As Boolean, CheckDoubles As Boolean, ResultLine As String, DecSec As String * 1
    MsgBox (StrComp(TypeName(dic_), "String"))
    IsDict = IIf(Not IsMissing(dic_) And StrComp(TypeName(dic_), "String") = 0 And (StrComp(dic_, "") <> 0), True, False)
    CheckDoubles = IIf(Not IsMissing(check_doubles), check_doubles, False) Or IIf(Not IsMissing(dic_) And StrComp(TypeName(dic_), "String") <> 0, CBool(dic_), False)
    DecSec = Application.DecimalSeparator
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True: objRegExp.IgnoreCase = True
    If (IsDict) Then
        objRegExp.Pattern = "[" & dic_ & "]*"
        ResultLine = objRegExp.Replace(line_, "")
    Else
        objRegExp.Pattern = "[^a-zA-Z0-9" & DecSec & "]*"
        ResultLine = objRegExp.Replace(line_, "")
    End If
    If (CheckDoubles) Then
        objRegExp.Pattern = "(.)(?=\1)"
        ResultLine = objRegExp.Replace(ResultLine, "")
    End If
    RepSymb = ResultLine
End Function