﻿#ExternalChecksum("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml","{406ea660-64cf-4c82-b6f0-42d48172a799}","516DE7D8FB18A16E584321E94B843A71")
'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports Fluent
Imports System
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Automation
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Ink
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Effects
Imports System.Windows.Media.Imaging
Imports System.Windows.Media.Media3D
Imports System.Windows.Media.TextFormatting
Imports System.Windows.Navigation
Imports System.Windows.Shapes
Imports System.Windows.Shell


'''<summary>
'''ShapeDialog
'''</summary>
<Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>  _
Partial Public Class ShapeDialog
    Inherits System.Windows.Window
    Implements System.Windows.Markup.IComponentConnector
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",4)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents CancelButton As Fluent.Button
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",9)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents OKButton As Fluent.Button
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",14)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents TypeComboBox As Fluent.ComboBox
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",15)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents SizeTextBox As Fluent.Spinner
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",16)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents TextBlock3 As System.Windows.Controls.TextBlock
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",17)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents ScrollViewer1 As System.Windows.Controls.ScrollViewer
    
    #End ExternalSource
    
    
    #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",18)
    <System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")>  _
    Friend WithEvents BorderSizeTextBox As Fluent.Spinner
    
    #End ExternalSource
    
    Private _contentLoaded As Boolean
    
    '''<summary>
    '''InitializeComponent
    '''</summary>
    <System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")>  _
    Public Sub InitializeComponent() Implements System.Windows.Markup.IComponentConnector.InitializeComponent
        If _contentLoaded Then
            Return
        End If
        _contentLoaded = true
        Dim resourceLocater As System.Uri = New System.Uri("/Document.Editor;component/dialogs/insert/shapedialog.xaml", System.UriKind.Relative)
        
        #ExternalSource("..\..\..\..\Dialogs\Insert\ShapeDialog.xaml",1)
        System.Windows.Application.LoadComponent(Me, resourceLocater)
        
        #End ExternalSource
    End Sub
    
    <System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0"),  _
     System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never),  _
     System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes"),  _
     System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity"),  _
     System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")>  _
    Sub System_Windows_Markup_IComponentConnector_Connect(ByVal connectionId As Integer, ByVal target As Object) Implements System.Windows.Markup.IComponentConnector.Connect
        If (connectionId = 1) Then
            Me.CancelButton = CType(target,Fluent.Button)
            Return
        End If
        If (connectionId = 2) Then
            Me.OKButton = CType(target,Fluent.Button)
            Return
        End If
        If (connectionId = 3) Then
            Me.TypeComboBox = CType(target,Fluent.ComboBox)
            Return
        End If
        If (connectionId = 4) Then
            Me.SizeTextBox = CType(target,Fluent.Spinner)
            Return
        End If
        If (connectionId = 5) Then
            Me.TextBlock3 = CType(target,System.Windows.Controls.TextBlock)
            Return
        End If
        If (connectionId = 6) Then
            Me.ScrollViewer1 = CType(target,System.Windows.Controls.ScrollViewer)
            Return
        End If
        If (connectionId = 7) Then
            Me.BorderSizeTextBox = CType(target,Fluent.Spinner)
            Return
        End If
        Me._contentLoaded = true
    End Sub
End Class

