Imports System, System.Text, System.CodeDom.Compiler, System.Reflection, System.IO, System.Diagnostics, Microsoft.VisualBasic
Public Class Plugins
    Private intCompilerErrors As CompilerErrorCollection

    Public Property CompilerErrors() As CompilerErrorCollection
        Get
            Return intCompilerErrors
        End Get
        Set(ByVal Value As CompilerErrorCollection)
            intCompilerErrors = Value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        intCompilerErrors = New CompilerErrorCollection
    End Sub

    Public Function Build(ByVal name As String, ByVal Code As String) As Object
        Dim codeProvider As New VBCodeProvider, parameters As New CompilerParameters, result As CompilerResults, assembly As System.Reflection.Assembly,
            instance As Object = Nothing, Plugin As Object = Nothing, Method As MethodInfo, type As Type
        Try
            ' Setup the Compiler Parameters  
            parameters.ReferencedAssemblies.Add("System.dll")
            parameters.ReferencedAssemblies.Add("System.XML.dll")
            parameters.ReferencedAssemblies.Add("System.Data.dll")
            parameters.ReferencedAssemblies.Add("System.Drawing.dll")
            parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            parameters.CompilerOptions = "/t:library"
            parameters.GenerateInMemory = True
            ' Generate the Code Framework
            Dim CodeSource As StringBuilder = New StringBuilder("")
            CodeSource.Append("Imports System, System.xml, System.Data, System.Drawing, System.Windows.Forms" & vbCrLf)
            CodeSource.Append("Namespace Plugins" & vbCrLf)
            CodeSource.Append(Code & vbCrLf)
            CodeSource.Append("End Namespace" & vbCrLf)
            Try
                ' Compile and get results 
                result = codeProvider.CompileAssemblyFromSource(parameters, CodeSource.ToString)
                If result.Errors.Count <> 0 Then
                    Me.CompilerErrors = result.Errors
                    Throw New Exception("Compile Errors")
                Else
                    assembly = result.CompiledAssembly
                    instance = assembly.CreateInstance("Plugins." & name)
                    type = instance.GetType
                    Method = type.GetMethod("StartPlugin")
                    Plugin = Method.Invoke(instance, Nothing)
                    Return Plugin
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                Stop
            End Try
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Stop
        End Try
        Return Plugin
    End Function

    Public Function IsStartupPlugin(ByVal name As String, ByVal Code As String) As Object
        Dim codeProvider As New VBCodeProvider, parameters As New CompilerParameters, result As CompilerResults, assembly As System.Reflection.Assembly,
            instance As Object = Nothing, Plugin As Object = Nothing, Method As MethodInfo, type As Type
        Try
            ' Setup the Compiler Parameters  
            parameters.ReferencedAssemblies.Add("System.dll")
            parameters.ReferencedAssemblies.Add("System.XML.dll")
            parameters.ReferencedAssemblies.Add("System.Data.dll")
            parameters.ReferencedAssemblies.Add("System.Drawing.dll")
            parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            parameters.CompilerOptions = "/t:library"
            parameters.GenerateInMemory = True
            ' Generate the Code Framework
            Dim CodeSource As StringBuilder = New StringBuilder("")
            CodeSource.Append("Imports System, System.xml, System.Data, System.Drawing, System.Windows.Forms" & vbCrLf)
            CodeSource.Append("Namespace Plugins" & vbCrLf)
            CodeSource.Append(Code & vbCrLf)
            CodeSource.Append("End Namespace" & vbCrLf)
            Try
                ' Compile and get results 
                result = codeProvider.CompileAssemblyFromSource(parameters, CodeSource.ToString)
                If result.Errors.Count <> 0 Then
                    Me.CompilerErrors = result.Errors
                    Throw New Exception("Compile Errors")
                Else
                    assembly = result.CompiledAssembly
                    instance = assembly.CreateInstance("Plugins." & name)
                    type = instance.GetType
                    Method = type.GetMethod("IsStartupPlugin")
                    Plugin = Method.Invoke(instance, Nothing)
                    Return Plugin
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                Stop
            End Try
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Stop
        End Try
        Return Plugin
    End Function

    Public Function IsEventPlugin(ByVal name As String, ByVal Code As String) As Object
        Dim codeProvider As New VBCodeProvider, parameters As New CompilerParameters, result As CompilerResults, assembly As System.Reflection.Assembly,
            instance As Object = Nothing, Plugin As Object = Nothing, Method As MethodInfo, type As Type
        Try
            ' Setup the Compiler Parameters  
            parameters.ReferencedAssemblies.Add("System.dll")
            parameters.ReferencedAssemblies.Add("System.XML.dll")
            parameters.ReferencedAssemblies.Add("System.Data.dll")
            parameters.ReferencedAssemblies.Add("System.Drawing.dll")
            parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            parameters.CompilerOptions = "/t:library"
            parameters.GenerateInMemory = True
            ' Generate the Code Framework
            Dim CodeSource As StringBuilder = New StringBuilder("")
            CodeSource.Append("Imports System, System.xml, System.Data, System.Drawing, System.Windows.Forms" & vbCrLf)
            CodeSource.Append("Namespace Plugins" & vbCrLf)
            CodeSource.Append(Code & vbCrLf)
            CodeSource.Append("End Namespace" & vbCrLf)
            Try
                ' Compile and get results 
                result = codeProvider.CompileAssemblyFromSource(parameters, CodeSource.ToString)
                If result.Errors.Count <> 0 Then
                    Me.CompilerErrors = result.Errors
                    Throw New Exception("Compile Errors")
                Else
                    assembly = result.CompiledAssembly
                    instance = assembly.CreateInstance("Plugins." & name)
                    type = instance.GetType
                    Method = type.GetMethod("IsEventPlugin")
                    Plugin = Method.Invoke(instance, Nothing)
                    Return Plugin
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                Stop
            End Try
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Stop
        End Try
        Return Plugin
    End Function
End Class