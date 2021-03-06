﻿using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class FunctionReturnValueNotUsedInspectionTests
    {
        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_ExplicitCallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Call Foo(""Test"")
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_CallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Foo ""Test""
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_AddressOf()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Bar AddressOf Foo
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_NoReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo() As Integer
End Function
Public Sub Bar()
    Foo
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_Ignored_DoesNotReturnResult_AddressOf()
        {
            const string inputCode =
@"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Bar AddressOf Foo
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_MultipleConsecutiveCalls()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    Foo Foo(Foo(""Bar""))
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IfStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    If Foo(""Test"") Then
    End If
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ForEachStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    For Each Bar In Foo
    Next Bar
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_WhileStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    While Foo
    Wend
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_DoUntilStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Do Until Foo
    Loop
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    TestVal = Foo(""Test"")
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_RecursiveFunction()
        {
            const string inputCode =
@"Public Function Factorial(ByVal n As Long) As Long
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(n - 1) * n
    End If
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ArgumentFunctionCall()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Bar Foo(""Test"")
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IgnoresBuiltInFunctions()
        {
            const string inputCode =
@"Public Sub Dummy()
    MsgBox ""Test""
    Workbooks.Add
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void GivenInterfaceImplementationMember_ReturnsNoResult()
        {
            const string interfaceCode =
@"Public Function Test() As Integer
End Function";

            const string implementationCode =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    Dim result As Integer
    result = testObj.Test
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("IFoo", ComponentType.ClassModule, interfaceCode);
            projectBuilder.AddComponent("Bar", ComponentType.ClassModule, implementationCode);
            projectBuilder.AddComponent("TestModule", ComponentType.StandardModule, callSiteCode);
            var vbe = projectBuilder.MockVbeBuilder().Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_InterfaceMember()
        {
            const string interfaceCode =
@"Public Function Test() As Integer
End Function";

            const string implementationCode =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Sub Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("IFoo", ComponentType.ClassModule, interfaceCode);
            projectBuilder.AddComponent("Bar", ComponentType.ClassModule, implementationCode);
            projectBuilder.AddComponent("TestModule", ComponentType.StandardModule, callSiteCode);
            var vbe = projectBuilder.MockVbeBuilder().Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
    If True Then
        Foo = _
        True
    Else
        Foo = False
    End If
End Function";

            const string expectedCode =
@"Public Sub Foo(ByVal bar As String)
    If True Then
    Else
    End If
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            string actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface_ManyBodyStatements()
        {
            const string inputCode =
@"Function foo(ByRef fizz As Boolean) As Boolean
    fizz = True
    goo
label1:
    foo = fizz
End Function

Sub goo()
End Sub";

            const string expectedCode =
@"Sub foo(ByRef fizz As Boolean)
    fizz = True
    goo
label1:
End Sub

Sub goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            string actual = module.Content();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks_Interface()
        {
            const string inputInterfaceCode =
@"Public Function Test() As Integer
End Function";

            const string expectedInterfaceCode =
@"Public Sub Test()
End Sub";

            const string inputImplementationCode1 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string inputImplementationCode2 =
@"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
@"
Public Function Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected);
            projectBuilder.AddComponent("IFoo", ComponentType.ClassModule, inputInterfaceCode);
            projectBuilder.AddComponent("Bar", ComponentType.ClassModule, inputImplementationCode1);
            projectBuilder.AddComponent("Bar2", ComponentType.ClassModule, inputImplementationCode2);
            projectBuilder.AddComponent("TestModule", ComponentType.StandardModule, callSiteCode);
            var vbe = projectBuilder.MockVbeBuilder().Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            var project = vbe.Object.VBProjects[0];
            var interfaceModule = project.VBComponents[0].CodeModule;
            string actualInterface = interfaceModule.Content();
            Assert.AreEqual(expectedInterfaceCode, actualInterface, "Interface");
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
@"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new FunctionReturnValueNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "FunctionReturnValueNotUsedInspection";
            var inspection = new FunctionReturnValueNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
