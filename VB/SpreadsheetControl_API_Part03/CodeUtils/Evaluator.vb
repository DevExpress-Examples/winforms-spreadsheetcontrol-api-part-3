Imports System
Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetControl_API_Part03

    Public MustInherit Class ExampleCodeEvaluator

        Protected MustOverride ReadOnly Property CodeStart As String

        Protected MustOverride ReadOnly Property CodeBeforeClasses As String

        Protected MustOverride ReadOnly Property CodeEnd As String

        Protected MustOverride Function GetCodeDomProvider() As CodeDomProvider

        Protected MustOverride Function GetModuleAssembly() As String

        Protected MustOverride Function GetExampleClassName() As String

        Public Function ExecuteCodeAndGenerateDocument(ByVal args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs) As Boolean
            Dim theCode As String = System.[String].Concat(Me.CodeStart, args.Code, Me.CodeBeforeClasses, args.CodeClasses, Me.CodeEnd)
            Dim linesOfCode As String() = New String() {theCode}
            Return Me.CompileAndRun(linesOfCode, args.EvaluationParameter)
        End Function

        Protected Friend Function CompileAndRun(ByVal linesOfCode As String(), ByVal evaluationParameter As Object) As Boolean
            Dim CompilerParams As System.CodeDom.Compiler.CompilerParameters = New System.CodeDom.Compiler.CompilerParameters()
            CompilerParams.GenerateInMemory = True
            CompilerParams.TreatWarningsAsErrors = False
            CompilerParams.GenerateExecutable = False
            Dim referencesSystem As String() = New String() {"System.dll", "System.Windows.Forms.dll", "System.Data.dll", "System.Xml.dll", "System.Core.dll", "System.Drawing.dll"}
            Dim referencesDX As String() = New String() {AssemblyInfo.SRAssemblyData, Me.GetModuleAssembly(), AssemblyInfo.SRAssemblyOfficeCore, AssemblyInfo.SRAssemblyPrintingCore, AssemblyInfo.SRAssemblyPrinting, AssemblyInfo.SRAssemblySpreadsheetCore, AssemblyInfo.SRAssemblyUtils}
            Dim references As String() = New String(referencesSystem.Length + referencesDX.Length - 1) {}
            For referenceIndex As Integer = 0 To referencesSystem.Length - 1
                references(referenceIndex) = referencesSystem(referenceIndex)
            Next

            Dim i As Integer = 0, initial As Integer = referencesSystem.Length
            While i < referencesDX.Length
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(referencesDX(i) & AssemblyInfo.FullAssemblyVersionExtension)
                If assembly IsNot Nothing Then references(i + initial) = assembly.Location
                i += 1
            End While

            CompilerParams.ReferencedAssemblies.AddRange(references)
            Dim provider As System.CodeDom.Compiler.CodeDomProvider = Me.GetCodeDomProvider()
            Dim compile As System.CodeDom.Compiler.CompilerResults = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode)
            If compile.Errors.HasErrors Then
                Dim text As String = "Compile error: "
                For Each ce As System.CodeDom.Compiler.CompilerError In compile.Errors
                    text += "rn" & ce.ToString()
                Next

                System.Windows.Forms.MessageBox.Show(text)
                Return False
            End If

            Dim [module] As System.Reflection.[Module] = Nothing
            Try
                [module] = compile.CompiledAssembly.GetModules()(0)
            Catch
            End Try

            Dim moduleType As System.Type = Nothing
            If [module] Is Nothing Then
                Return False
            End If

            moduleType = [module].[GetType](Me.GetExampleClassName())
            Dim methInfo As System.Reflection.MethodInfo = Nothing
            If moduleType Is Nothing Then
                Return False
            End If

            methInfo = moduleType.GetMethod("Process")
            If methInfo IsNot Nothing Then
                Try
                    methInfo.Invoke(Nothing, New Object() {evaluationParameter})
                Catch __unusedException1__ As System.Exception
                    Return False ' an error
                End Try

                Return True
            End If

            Return False
        End Function
    End Class

    Public MustInherit Class RichEditExampleCodeEvaluator
        Inherits SpreadsheetControl_API_Part03.ExampleCodeEvaluator

        Protected Overrides Function GetModuleAssembly() As String
            Return AssemblyInfo.SRAssemblyRichEditCore
        End Function

        Protected Overrides Function GetExampleClassName() As String
            Return "RichEditCodeResultViewer.ExampleItem"
        End Function
    End Class

#Region "RichEditCSExampleCodeEvaluator"
    Public Class RichEditCSExampleCodeEvaluator
        Inherits SpreadsheetControl_API_Part03.RichEditExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New Microsoft.CSharp.CSharpCodeProvider()
        End Function

        Const codeStartField As String = "using System;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Spreadsheet;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Drawing;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Windows.Forms;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using DevExpress.Utils;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.IO;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Diagnostics;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Xml;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Data;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Collections.Generic;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Linq;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "using System.Globalization;" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "namespace RichEditCodeResultViewer { " & Global.Microsoft.VisualBasic.Constants.vbCrLf & "public class ExampleItem { " & Global.Microsoft.VisualBasic.Constants.vbCrLf & "        public static void Process(IWorkbook workbook) { " & Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Const codeBeforeClassesField As String = "       " & Global.Microsoft.VisualBasic.Constants.vbCrLf & " }" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "    }" & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Const codeEndField As String = Global.Microsoft.VisualBasic.Constants.vbCrLf & "    }" & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Protected Overrides ReadOnly Property CodeStart As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditCSExampleCodeEvaluator.codeStartField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeBeforeClasses As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditCSExampleCodeEvaluator.codeBeforeClassesField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeEnd As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditCSExampleCodeEvaluator.codeEndField
            End Get
        End Property
    End Class

#End Region
#Region "RichEditVbExampleCodeEvaluator"
    Public Class RichEditVbExampleCodeEvaluator
        Inherits SpreadsheetControl_API_Part03.RichEditExampleCodeEvaluator

        Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
            Return New Microsoft.VisualBasic.VBCodeProvider()
        End Function

        Const codeStartField As String = "Imports Microsoft.VisualBasic" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Spreadsheet" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Drawing" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Windows.Forms" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports DevExpress.Utils" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.IO" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Diagnostics" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Xml" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Data" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Collections.Generic" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Imports System.Globalization" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "Namespace RichEditCodeResultViewer" & Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbTab & "Public Class ExampleItem" & Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbTab & Global.Microsoft.VisualBasic.Constants.vbTab & "Public Shared Sub Process(ByVal workbook As IWorkbook)" & Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Const codeBeforeClassesField As String = Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbTab & Global.Microsoft.VisualBasic.Constants.vbTab & "End Sub" & Global.Microsoft.VisualBasic.Constants.vbCrLf & Global.Microsoft.VisualBasic.Constants.vbTab & "End Class" & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Const codeEndField As String = Global.Microsoft.VisualBasic.Constants.vbCrLf & "End Namespace" & Global.Microsoft.VisualBasic.Constants.vbCrLf

        Protected Overrides ReadOnly Property CodeStart As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditVbExampleCodeEvaluator.codeStartField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeBeforeClasses As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditVbExampleCodeEvaluator.codeBeforeClassesField
            End Get
        End Property

        Protected Overrides ReadOnly Property CodeEnd As String
            Get
                Return SpreadsheetControl_API_Part03.RichEditVbExampleCodeEvaluator.codeEndField
            End Get
        End Property
    End Class

#End Region
    Public MustInherit Class ExampleEvaluatorByTimer
        Implements System.IDisposable

        Private leakSafeCompileEventRouter As SpreadsheetControl_API_Part03.LeakSafeCompileEventRouter

        Private compileExampleTimer As System.Windows.Forms.Timer

        Private compileComplete As Boolean = True

        Const CompileTimeIntervalInMilliseconds As Integer = 2000

        Public Sub New(ByVal enableTimer As Boolean)
            Me.leakSafeCompileEventRouter = New SpreadsheetControl_API_Part03.LeakSafeCompileEventRouter(Me)
            If enableTimer Then
                Me.compileExampleTimer = New System.Windows.Forms.Timer()
                Me.compileExampleTimer.Interval = SpreadsheetControl_API_Part03.ExampleEvaluatorByTimer.CompileTimeIntervalInMilliseconds
                AddHandler Me.compileExampleTimer.Tick, New System.EventHandler(AddressOf Me.leakSafeCompileEventRouter.OnCompileExampleTimerTick) 'OnCompileTimerTick
                Me.compileExampleTimer.Enabled = True
            End If
        End Sub

        Public Sub New()
            Me.New(True)
        End Sub

#Region "Events"
        Public Event QueryEvaluate As SpreadsheetControl_API_Part03.CodeEvaluationEventHandler

        Protected Friend Overridable Function RaiseQueryEvaluate() As CodeEvaluationEventArgs
            If QueryEvaluateEvent IsNot Nothing Then
                Dim args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs = New SpreadsheetControl_API_Part03.CodeEvaluationEventArgs()
                RaiseEvent QueryEvaluate(Me, args)
                Return args
            End If

            Return Nothing
        End Function

        Public Event OnBeforeCompile As System.EventHandler

        Private Sub RaiseOnBeforeCompile()
            RaiseEvent OnBeforeCompile(Me, New System.EventArgs())
        End Sub

        Public Event OnAfterCompile As SpreadsheetControl_API_Part03.OnAfterCompileEventHandler

        Private Sub RaiseOnAfterCompile(ByVal result As Boolean)
            RaiseEvent OnAfterCompile(Me, New SpreadsheetControl_API_Part03.OnAfterCompileEventArgs() With {.Result = result})
        End Sub

#End Region
        Public Sub CompileExample(ByVal sender As Object, ByVal e As System.EventArgs)
            If Not Me.compileComplete Then Return
            Dim args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs = Me.RaiseQueryEvaluate()
            If Not args.Result Then Return
            Me.ForceCompile(args)
        End Sub

        Public Sub ForceCompile(ByVal args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs)
            Me.compileComplete = False
            If Not System.[String].IsNullOrEmpty(args.Code) Then Me.CompileExampleAndShowPrintPreview(args)
            Me.compileComplete = True
        End Sub

        Private Sub CompileExampleAndShowPrintPreview(ByVal args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs)
            Dim evaluationSucceed As Boolean = False
            Try
                Me.RaiseOnBeforeCompile()
                evaluationSucceed = Me.Evaluate(args)
            Finally
                Me.RaiseOnAfterCompile(evaluationSucceed)
            End Try
        End Sub

        Public Function Evaluate(ByVal args As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs) As Boolean
            Dim richeditExampleCodeEvaluator As SpreadsheetControl_API_Part03.ExampleCodeEvaluator = Me.GetExampleCodeEvaluator(args.Language)
            Return richeditExampleCodeEvaluator.ExecuteCodeAndGenerateDocument(args)
        End Function

        Protected MustOverride Function GetExampleCodeEvaluator(ByVal language As SpreadsheetControl_API_Part03.ExampleLanguage) As ExampleCodeEvaluator

        Public Sub Dispose() Implements Global.System.IDisposable.Dispose
            If Me.compileExampleTimer IsNot Nothing Then
                Me.compileExampleTimer.Enabled = False
                If Me.leakSafeCompileEventRouter IsNot Nothing Then RemoveHandler Me.compileExampleTimer.Tick, New System.EventHandler(AddressOf Me.leakSafeCompileEventRouter.OnCompileExampleTimerTick) 'OnCompileTimerTick
                Me.compileExampleTimer.Dispose()
                Me.compileExampleTimer = Nothing
            End If
        End Sub
    End Class

#Region "RichEditExampleEvaluatorByTimer"
    Public Class RichEditExampleEvaluatorByTimer
        Inherits SpreadsheetControl_API_Part03.ExampleEvaluatorByTimer

        Public Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Function GetExampleCodeEvaluator(ByVal language As SpreadsheetControl_API_Part03.ExampleLanguage) As ExampleCodeEvaluator
            If language = SpreadsheetControl_API_Part03.ExampleLanguage.VB Then Return New SpreadsheetControl_API_Part03.RichEditVbExampleCodeEvaluator()
            Return New SpreadsheetControl_API_Part03.RichEditCSExampleCodeEvaluator()
        End Function
    End Class

#End Region
#Region "LeakSafeCompileEventRouter"
    Public Class LeakSafeCompileEventRouter

        Private ReadOnly weakControlRef As System.WeakReference

        Public Sub New(ByVal [module] As SpreadsheetControl_API_Part03.ExampleEvaluatorByTimer)
            'Guard.ArgumentNotNull(module, "module");
            Me.weakControlRef = New System.WeakReference([module])
        End Sub

        Public Sub OnCompileExampleTimerTick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim [module] As SpreadsheetControl_API_Part03.ExampleEvaluatorByTimer = CType(Me.weakControlRef.Target, SpreadsheetControl_API_Part03.ExampleEvaluatorByTimer)
            If [module] IsNot Nothing Then [module].CompileExample(sender, e)
        End Sub
    End Class

    Public Class CodeEvaluationEventArgs
        Inherits System.EventArgs

        Public Property Result As Boolean

        Public Property Code As String

        Public Property CodeClasses As String

        Public Property Language As ExampleLanguage

        Public Property EvaluationParameter As Object
    End Class

    Public Delegate Sub CodeEvaluationEventHandler(ByVal sender As Object, ByVal e As SpreadsheetControl_API_Part03.CodeEvaluationEventArgs)

    Public Class OnAfterCompileEventArgs
        Inherits System.EventArgs

        Public Property Result As Boolean
    End Class

    Public Delegate Sub OnAfterCompileEventHandler(ByVal sender As Object, ByVal e As SpreadsheetControl_API_Part03.OnAfterCompileEventArgs)
#End Region
End Namespace
