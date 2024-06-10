If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUISERVER")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 225,31,false
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00016"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00016"
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "4016185"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
session.findById("wnd[1]/usr/txtRV45A-POSNR").text = "20"
session.findById("wnd[1]/usr/txtRV45A-POSNR").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,0]").key = "CB"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/tblSAPMV45ATCTRL_UPOS_ABSAGE/cmbVBAP-ABGRU[2,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
