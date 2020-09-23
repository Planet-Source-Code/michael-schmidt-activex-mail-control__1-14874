Attribute VB_Name = "modRegistry"
Option Explicit

'============================================
'   Registry - Load Settings
'============================================
Public Sub LoadAppSettings()

    ' Load User Settings
    frmSample.txtPassword = GetSetting(App.Title, "SETTINGS", "PASSWORD")
    frmSample.txtUser = GetSetting(App.Title, "SETTINGS", "USER")
    
    ' Load Mail Settings
    frmSample.txtDelay = GetSetting(App.Title, "SETTINGS", "DELAY", "2")
    frmSample.txtServer = GetSetting(App.Title, "SETTINGS", "SERVER")

End Sub


'============================================
'   Registry - Save Settings
'============================================
Public Sub SaveAppSettings()

    ' Load User Settings
    SaveSetting App.Title, "SETTINGS", "PASSWORD", frmSample.txtPassword
    SaveSetting App.Title, "SETTINGS", "USER", frmSample.txtUser
    
    ' Load Mail Settings
    SaveSetting App.Title, "SETTINGS", "DELAY", frmSample.txtDelay
    SaveSetting App.Title, "SETTINGS", "SERVER", frmSample.txtServer

End Sub


'============================================
'   Registry - Delete Settings
'============================================
Public Sub DeleteAppSettings()
On Error Resume Next
    ' Delete Application Settings
    DeleteSetting App.Title, "SETTINGS"


End Sub
