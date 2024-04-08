VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdateEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements LSEXT.IEntityExtension

Private Function IEntityExtension_CanExecute(Parameters As LSEXT.IExtensionParameters) As LSEXT.ExecuteExtension
    IEntityExtension_CanExecute = exEnabled
End Function

Private Sub IEntityExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
    Dim ServiceProvider As LSSERVICEPROVIDERLib.NautilusServiceProvider
    Dim NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
    Dim con As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim constr As String
    
    Set rst = Parameters("RECORDS")
    
    Set ServiceProvider = Parameters("SERVICE_PROVIDER")
    Set NtlsCon = ServiceProvider.QueryServiceProvider("DBConnection")
    constr = "Provider=OraOLEDB.Oracle" & _
        ";Data Source=" & NtlsCon.GetServerDetails & _
        ";User ID=" & NtlsCon.GetUsername & _
        ";Password=" & NtlsCon.GetPassword
        constr = NtlsCon.GetADOConnectionString
        
        
    con.Open constr
    con.CursorLocation = adUseClient
    con.Execute "SET ROLE LIMS_USER"
    Call ConnectSameSession(con, CDbl(Parameters("SESSION_ID")))
   
    con.Execute ("call lims.workflow_events.UpdateEvents(" & rst("WORKFLOW_ID") & ")")
    MsgBox "Events field for objects referenced by this workflow was updated successfully!"
End Sub

Private Sub ConnectSameSession(ByVal aConnection As ADODB.Connection, ByVal aSessionID As Double)
    Dim aProc As New ADODB.Command
    Dim aSession As New ADODB.Parameter
    
    aProc.ActiveConnection = aConnection
    aProc.CommandText = "lims.lims_env.connect_same_session"
    aProc.CommandType = adCmdStoredProc
    
    aSession.Type = adDouble
    aSession.Direction = adParamInput
    aSession.Value = aSessionID
    aProc.Parameters.Append aSession
    
    aProc.Execute
    Set aSession = Nothing
    Set aProc = Nothing
End Sub

Private Function nvl(e As Variant, v As Variant) As Variant
    nvl = IIf(IsNull(e), v, e)
End Function


