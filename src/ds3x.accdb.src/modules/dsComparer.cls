VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0

Implements mscorlib.IComparer


'public class SimpleStringComparer : IComparer
'{
'    int IComparer.Compare(object x, object y)
'    {
'        string cmpstr = (string)x;
'        return cmpstr.CompareTo((string)y);
'    }
'}


Private Function IComparer_Compare(ByVal X As Variant, ByVal Y As Variant) As Long
    IComparer_Compare = CLng(CLng(X) - CLng(Y))
End Function




