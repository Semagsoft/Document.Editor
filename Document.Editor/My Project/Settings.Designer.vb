﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On



<Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
 Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.3.0.0"),  _
 Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
Partial Friend NotInheritable Class MySettings
    Inherits Global.System.Configuration.ApplicationSettingsBase
    
    Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
    
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
    
    Public Shared ReadOnly Property [Default]() As MySettings
        Get
            
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
            Return defaultInstance
        End Get
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("4")>  _
    Public Property Options_Theme() As Integer
        Get
            Return CType(Me("Options_Theme"),Integer)
        End Get
        Set
            Me("Options_Theme") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property MainWindow_IsMax() As Boolean
        Get
            Return CType(Me("MainWindow_IsMax"),Boolean)
        End Get
        Set
            Me("MainWindow_IsMax") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_StartupMode() As Integer
        Get
            Return CType(Me("Options_StartupMode"),Integer)
        End Get
        Set
            Me("Options_StartupMode") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("1")>  _
    Public Property Options_Tabs_SizeMode() As Integer
        Get
            Return CType(Me("Options_Tabs_SizeMode"),Integer)
        End Get
        Set
            Me("Options_Tabs_SizeMode") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property Options_SpellCheck() As Boolean
        Get
            Return CType(Me("Options_SpellCheck"),Boolean)
        End Get
        Set
            Me("Options_SpellCheck") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("<?xml version=""1.0"" encoding=""utf-16""?>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"<ArrayOfString xmlns:xsi=""http://www.w3."& _ 
        "org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" />")>  _
    Public Property Options_RecentFiles() As Global.System.Collections.Specialized.StringCollection
        Get
            Return CType(Me("Options_RecentFiles"),Global.System.Collections.Specialized.StringCollection)
        End Get
        Set
            Me("Options_RecentFiles") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_Tabs_CloseButtonMode() As Integer
        Get
            Return CType(Me("Options_Tabs_CloseButtonMode"),Integer)
        End Get
        Set
            Me("Options_Tabs_CloseButtonMode") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_TTSV() As Integer
        Get
            Return CType(Me("Options_TTSV"),Integer)
        End Get
        Set
            Me("Options_TTSV") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("en-US")>  _
    Public Property Options_Language() As String
        Get
            Return CType(Me("Options_Language"),String)
        End Get
        Set
            Me("Options_Language") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property MainWindow_ShowStatusBar() As Boolean
        Get
            Return CType(Me("MainWindow_ShowStatusBar"),Boolean)
        End Get
        Set
            Me("MainWindow_ShowStatusBar") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_TTSS() As Integer
        Get
            Return CType(Me("Options_TTSS"),Integer)
        End Get
        Set
            Me("Options_TTSS") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property SaveDialog_IsMax() As Boolean
        Get
            Return CType(Me("SaveDialog_IsMax"),Boolean)
        End Get
        Set
            Me("SaveDialog_IsMax") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property Options_ShowRecentDocuments() As Boolean
        Get
            Return CType(Me("Options_ShowRecentDocuments"),Boolean)
        End Get
        Set
            Me("Options_ShowRecentDocuments") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property App_FirstRun() As Boolean
        Get
            Return CType(Me("App_FirstRun"),Boolean)
        End Get
        Set
            Me("App_FirstRun") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property Options_EnablePlugins() As Boolean
        Get
            Return CType(Me("Options_EnablePlugins"),Boolean)
        End Get
        Set
            Me("Options_EnablePlugins") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("<?xml version=""1.0"" encoding=""utf-16""?>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"<ArrayOfString xmlns:xsi=""http://www.w3."& _ 
        "org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  <s"& _ 
        "tring>&lt;?xml version=""1.0"" encoding=""utf-16""?&gt;</string>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  <string>&lt;Arra"& _ 
        "yOfString xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http:"& _ 
        "//www.w3.org/2001/XMLSchema"" /&gt;</string>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"</ArrayOfString>")>  _
    Public Property Options_PreviousDocuments() As Global.System.Collections.Specialized.StringCollection
        Get
            Return CType(Me("Options_PreviousDocuments"),Global.System.Collections.Specialized.StringCollection)
        End Get
        Set
            Me("Options_PreviousDocuments") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property Options_ShowStartupDialog() As Boolean
        Get
            Return CType(Me("Options_ShowStartupDialog"),Boolean)
        End Get
        Set
            Me("Options_ShowStartupDialog") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property Options_EnableGlass() As Boolean
        Get
            Return CType(Me("Options_EnableGlass"),Boolean)
        End Get
        Set
            Me("Options_EnableGlass") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("Calibri")>  _
    Public Property Options_DefaultFont() As Global.System.Windows.Media.FontFamily
        Get
            Return CType(Me("Options_DefaultFont"),Global.System.Windows.Media.FontFamily)
        End Get
        Set
            Me("Options_DefaultFont") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("14")>  _
    Public Property Options_DefaultFontSize() As Double
        Get
            Return CType(Me("Options_DefaultFontSize"),Double)
        End Get
        Set
            Me("Options_DefaultFontSize") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property MainWindow_ShowRuler() As Boolean
        Get
            Return CType(Me("MainWindow_ShowRuler"),Boolean)
        End Get
        Set
            Me("MainWindow_ShowRuler") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_RulerMeasurement() As Integer
        Get
            Return CType(Me("Options_RulerMeasurement"),Integer)
        End Get
        Set
            Me("Options_RulerMeasurement") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property MainWindow_ShowHomeTab() As Boolean
        Get
            Return CType(Me("MainWindow_ShowHomeTab"),Boolean)
        End Get
        Set
            Me("MainWindow_ShowHomeTab") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property ShowCommonEditGroup() As Boolean
        Get
            Return CType(Me("ShowCommonEditGroup"),Boolean)
        End Get
        Set
            Me("ShowCommonEditGroup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property ShowCommonViewGroup() As Boolean
        Get
            Return CType(Me("ShowCommonViewGroup"),Boolean)
        End Get
        Set
            Me("ShowCommonViewGroup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
    Public Property ShowCommonInsertGroup() As Boolean
        Get
            Return CType(Me("ShowCommonInsertGroup"),Boolean)
        End Get
        Set
            Me("ShowCommonInsertGroup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property ShowCommonFormatGroup() As Boolean
        Get
            Return CType(Me("ShowCommonFormatGroup"),Boolean)
        End Get
        Set
            Me("ShowCommonFormatGroup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property ShowCommonToolsGroup() As Boolean
        Get
            Return CType(Me("ShowCommonToolsGroup"),Boolean)
        End Get
        Set
            Me("ShowCommonToolsGroup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Options_TabPlacement() As Integer
        Get
            Return CType(Me("Options_TabPlacement"),Integer)
        End Get
        Set
            Me("Options_TabPlacement") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("48")>  _
    Public Property MainWindow_Left() As Double
        Get
            Return CType(Me("MainWindow_Left"),Double)
        End Get
        Set
            Me("MainWindow_Left") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("48")>  _
    Public Property MainWindow_Top() As Double
        Get
            Return CType(Me("MainWindow_Top"),Double)
        End Get
        Set
            Me("MainWindow_Top") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("668")>  _
    Public Property MainWindow_Height() As Double
        Get
            Return CType(Me("MainWindow_Height"),Double)
        End Get
        Set
            Me("MainWindow_Height") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("968")>  _
    Public Property MainWindow_Width() As Double
        Get
            Return CType(Me("MainWindow_Width"),Double)
        End Get
        Set
            Me("MainWindow_Width") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
    Public Property Options_CheckForUpdatesOnStartup() As Boolean
        Get
            Return CType(Me("Options_CheckForUpdatesOnStartup"),Boolean)
        End Get
        Set
            Me("Options_CheckForUpdatesOnStartup") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("")>  _
    Public Property Options_TemplatesFolder() As String
        Get
            Return CType(Me("Options_TemplatesFolder"),String)
        End Get
        Set
            Me("Options_TemplatesFolder") = value
        End Set
    End Property
End Class

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.Document.Editor.MySettings
            Get
                Return Global.Document.Editor.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
