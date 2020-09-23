Attribute VB_Name = "modEmail"
Type email
    Subject As String
    To As String
    From As String
    Msg As String
    Format As String
    SMTP As String 'SMTP SERVER
End Type

Public Myemail As email
