Imports System.Windows.Forms
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Systems.Middle

Public Class Form1
    Private WithEvents m_oEVSelectSet As SelectSet
    Private oSelect As SelectSet = ClientServiceProvider.SelectSet
    Private m_UOM As UOMManager
    Public Function MyAngle(A As Vector, B As Vector) As Double
        Dim angle As Double
        angle = Math.Acos(A.Dot(B) / A.Length * B.Length)
        Return angle
    End Function
    Public Function MyRotate(E As Vector, R As Vector, angle As Double) As Vector
        E.Length = 1
        Dim A As Vector
        A = New Vector（R.X * Math.Cos(angle), R.Y * Math.Cos(angle), R.Z * Math.Cos(angle))
        Dim B As Vector
        B = E.Cross(R)
        Dim C As Vector
        C = New Vector(B.X * Math.Sin(angle), B.Y * Math.Sin(angle), B.Z * Math.Sin(angle))
        Dim R1 As Vector
        R1 = A + C
        Return R1
    End Function

    Public Function TapOD(OD As Double, ByRef NS As Double, ByRef DIN As Double) As String
        Dim TapNum As String = Nothing
        Select Case OD
            Case 2 To 2.05
                NS = 80
                DIN = 2000
                TapNum = "Tap-084"
            Case 1.9 To 1.95
                NS = 76
                DIN = 1900
                TapNum = "Tap-083"
            Case 1.8 To 1.85
                NS = 72
                DIN = 1800
                TapNum = "Tap-082"
            Case 1.7 To 1.75
                NS = 68
                DIN = 1700
                TapNum = "Tap-081"
            Case 1.6 To 1.65
                NS = 64
                DIN = 1600
                TapNum = "Tap-080"
            Case 1.5 To 1.55
                NS = 60
                DIN = 1500
                TapNum = "Tap-079"
            Case 1.4 To 1.45
                NS = 56
                DIN = 1400
                TapNum = "Tap-078"
            Case 1.3 To 1.35
                NS = 52
                DIN = 1300
                TapNum = "Tap-077"
            Case 1.2 To 1.25
                NS = 48
                DIN = 1200
                TapNum = "Tap-037"
            Case 1.15 To 1.2
                NS = 46
                DIN = 1150
                TapNum = "Tap-076"
            Case 1.1 To 1.15
                NS = 44
                DIN = 1100
                TapNum = "Tap-075"
            Case 1.05 To 1.1
                NS = 42
                DIN = 1050
                TapNum = "Tap-036"
            Case 1.0 To 1.05
                NS = 40
                DIN = 1000
                TapNum = "Tap-074"
            Case 0.95 To 1
                NS = 38
                DIN = 950
                TapNum = "Tap-073"
            Case 0.9 To 0.95
                NS = 36
                DIN = 900
                TapNum = "Tap-035"
            Case 0.85 To 0.9
                NS = 34
                DIN = 850
                TapNum = "Tap-034"
            Case 0.8 To 0.85
                NS = 32
                DIN = 800
                TapNum = "Tap-033"
            Case 0.75 To 0.8
                NS = 30
                DIN = 750
                TapNum = "Tap-032"
            Case 0.7 To 0.75
                NS = 28
                DIN = 700
                TapNum = "Tap-031"
            Case 0.65 To 0.7
                NS = 26
                DIN = 650
                TapNum = "Tap-030"
            Case 0.6 To 0.65
                NS = 24
                DIN = 600
                TapNum = "Tap-029"
            Case 0.55 To 0.6
                NS = 22
                DIN = 550
                TapNum = "Tap-072"
            Case 0.5 To 0.55
                NS = 20
                DIN = 500
                TapNum = "Tap-028"
            Case 0.45 To 0.5
                NS = 18
                DIN = 450
                TapNum = "Tap-027"
            Case 0.4 To 0.45
                NS = 16
                DIN = 400
                TapNum = "Tap-026"
            Case 0.35 To 0.4
                NS = 14
                DIN = 350
                TapNum = "Tap-025"
            Case 0.32 To 0.35
                NS = 12
                DIN = 300
                TapNum = "Tap-024"
            Case 0.27 To 0.32
                NS = 10
                DIN = 250
                TapNum = "Tap-023"
            Case 0.21 To 0.27
                NS = 8
                DIN = 200
                TapNum = "Tap-022"
            Case 0.16 To 0.21
                NS = 6
                DIN = 150
                TapNum = "Tap-021"
            Case 0.14 To 0.16
                NS = 5
                DIN = 125
                TapNum = "Tap-020"
            Case 0.11 To 0.14
                NS = 4
                DIN = 100
                TapNum = "Tap-019"
            Case 0.1 To 0.11
                NS = 3.5
                DIN = 90
                TapNum = "Tap-017"
            Case 0.085 To 0.1
                NS = 3
                DIN = 80
                TapNum = "Tap-016"
            Case 0.075 To 0.085
                NS = 2.5
                DIN = 65
                TapNum = "Tap-014"
            Case 0.06 To 0.075
                NS = 2
                DIN = 50
                TapNum = "Tap-013"
            Case 0.048 To 0.06
                NS = 1.5
                DIN = 40
                TapNum = "Tap-011"
            Case 0.042 To 0.048
                NS = 1.25
                DIN = 32
                TapNum = "Tap-009"
            Case 0.033 To 0.042
                NS = 1
                DIN = 25
                TapNum = "Tap-007"
            Case 0.026 To 0.033
                NS = 0.75
                DIN = 20
                TapNum = "Tap-005"
            Case 0.021 To 0.026
                NS = 0.5
                DIN = 15
                TapNum = "Tap-003"
            Case 0.017 To 0.021
                NS = 0.375
                DIN = 10
                TapNum = "Tap-001"
            Case 0.013 To 0.017
                NS = 0.25
                DIN = 8
                TapNum = "Tap-071"
            Case 0.01 To 0.013
                NS = 0.125
                DIN = 6
                TapNum = "Tap-070"
        End Select
        Return TapNum
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Text = My.Settings.text
        TextBox1.Text = My.Settings.note
        If My.Settings.ReportingRequirement = 2 Then
            SwitchButton1.Value = False
        Else
            SwitchButton1.Value = True
        End If
        m_oEVSelectSet = oSelect
        If m_oEVSelectSet.SelectedObjects.Count = 1 AndAlso TypeOf m_oEVSelectSet.SelectedObjects.Item(0) Is Support Then
            ToolStripStatusLabel1.Text = m_oEVSelectSet.SelectedObjects.Item(0).GetPropertyValue("IJNamedItem", "Name").ToString
        End If
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyValue = 27 Then
            Close()
        ElseIf e.KeyCode = Keys.F1 AndAlso e.Modifiers = Keys.Control Then
            If ButtonX1.Visible Then
                Text = "jiazheng20230627"
                ButtonX1.Visible = False
                ButtonX2.Visible = False
                TextBox1.Visible = True
                SwitchButton1.Visible = True
            Else
                Text = My.Settings.text
                ButtonX1.Visible = True
                ButtonX2.Visible = True
                TextBox1.Visible = False
                SwitchButton1.Visible = False
            End If
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        My.Settings.note = TextBox1.Text
    End Sub

    Private Sub SwitchButton1_ValueChanged(sender As Object, e As EventArgs) Handles SwitchButton1.ValueChanged
        If SwitchButton1.Value = True Then
            My.Settings.ReportingRequirement = 5
            My.Settings.ReportingType = 5
        Else
            My.Settings.ReportingRequirement = 2
            My.Settings.ReportingType = 2
        End If
    End Sub

    Private Sub m_oEVSelectSet_SelectionChanged(eventSpecifics As MultiCasterEventArgs) Handles m_oEVSelectSet.SelectionChanged
        If m_oEVSelectSet.SelectedObjects.Count = 1 AndAlso TypeOf m_oEVSelectSet.SelectedObjects.Item(0) Is Support Then
            ToolStripStatusLabel1.Text = m_oEVSelectSet.SelectedObjects.Item(0).GetPropertyValue("IJNamedItem", "Name").ToString
        Else
            ToolStripStatusLabel1.Text = ""
        End If
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ClientServiceProvider.TransactionMgr.Abort()
        m_oEVSelectSet = Nothing
    End Sub

    Private Sub ButtonX2_Click(sender As Object, e As EventArgs) Handles ButtonX2.Click
        m_UOM = MiddleServiceProvider.UOMMgr
        If oSelect.SelectedObjects.Count = 1 Then
            If TypeOf (oSelect.SelectedObjects.Item(0)) Is Support Then
                Try
                    Dim oSupport As Support = oSelect.SelectedObjects.Item(0)
                    Dim oPipe As RouteFeature = oSupport.SupportedObjects.Item(0)
                    Dim run1 As PipeRun = oPipe.SystemParent
                    Dim oPipeTapFeature As PipeTapFeature
                    Dim Batch As PipeBranchFeature
                    Dim run2 As PipeRun
                    Dim position1 As Position = New Position(oSupport.SupportPosition)
                    Dim OPD As PropertyValueDouble = Nothing
                    Dim oRefport As RefPortHelper = New RefPortHelper(oSupport)
                    Dim oMat As Matrix4X4 = oRefport.PortLCS("Route")
                    Dim oVec As Vector = oMat.ZAxis
                    Dim oVecX As Vector = oMat.XAxis
                    Dim vector2 As Vector, vector3 As Vector
                    Dim vector4 As Vector = New Vector()
                    Dim OD As Double
                    Dim Offset1 As Double
                    Dim vector6 As Vector
                    Dim angle1 As Double = 0
                    For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                        If Dummy.BOMDescription.Contains("PIPE") Then
                            OPD = Dummy.GetPropertyValue("IJOAhsPipeOD", "PipeOD")
                        End If
                    Next
                    If OPD Is Nothing Then
                        ToolStripStatusLabel1.Text = "不是耳轴"
                        Exit Sub
                    Else
                        OD = m_UOM.ParseUnit(UnitType.Distance, OPD.ToString)
                    End If
                    If oPipe.ClassInfo.Name = "CPPipeStraightPathFeat" Then
                        vector6 = oPipe.EndLocation - oPipe.StartLocation
                        For Each oAlong As AlongLegFeature In oPipe.Parts.Item(0).GetRelationship("PathGeneratedParts", "DefiningFeature").TargetObjects.OfType(Of AlongLegFeature)
                            oPipeTapFeature = Nothing
                            Try
                                oPipeTapFeature = oAlong.GetRelationship("HasTapFeatures", "TapFeature").TargetObjects.Item(0)
                                Batch = oAlong.GetRelationship("OffLineFeatures", "OffLineFeature").TargetObjects.Item(0)
                            Catch ex As Exception
                                Batch = Nothing
                            End Try
                            If Batch IsNot Nothing Then
                                run2 = Batch.GetRelationship("PathSpecification", "thePathSystemInfo").TargetObjects.Item(0)
                                If oSupport.Name.Contains("S030") Or oSupport.Name.Contains("S031") Or oSupport.Name.Contains("S264") Then
                                    Offset1 = m_UOM.ParseUnit(UnitType.Distance, oSupport.GetPropertyValue("IJUAhsOffset1", "Offset1").ToString)
                                    vector6.Length = OD / 2 + Offset1
                                    If oVecX.Add(vector6).Length < oVecX.Subtract(vector6).Length Then
                                        vector6.Length = -vector6.Length
                                    End If
                                    position1.Set(vector6.X + oSupport.SupportPosition.X, vector6.Y + oSupport.SupportPosition.Y, vector6.Z + oSupport.SupportPosition.Z)
                                    For Each oFeat As PipeStraightFeature In run2.Features.OfType(Of PipeStraightFeature)
                                        If oFeat.Parts.Item(0).Name = "Dummy Pipe" And position1 = oAlong.Location Then
                                            run2.Delete()
                                            ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                            oPipeTapFeature.Delete()
                                            ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                        End If
                                    Next
                                Else
                                    For Each oFeat As PipeStraightFeature In run2.Features.OfType(Of PipeStraightFeature)
                                        If oFeat.Parts.Item(0).Name = "Dummy Pipe" And oSupport.SupportPosition = oAlong.Location Then
                                            run2.Delete()
                                            ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                            oPipeTapFeature.Delete()
                                            ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    ElseIf oPipe.ClassInfo.Name = "CPPipeTurnPathFeat" Then
                        vector2 = oPipe.Location - oPipe.StartLocation
                        vector3 = oPipe.Location - oPipe.EndLocation
                        vector6 = vector2.Cross(vector3)
                        Dim d1 As Double = Math.Abs(oSupport.SupportPosition.DistanceToPoint(oPipe.StartLocation))
                        Dim d2 As Double = Math.Abs(oSupport.SupportPosition.DistanceToPoint(oPipe.EndLocation))
                        angle1 = m_UOM.ParseUnit(UnitType.Angle, oSupport.GetRelationship("AssemblyHierarchy", "IsChildOf").TargetObjects(0).GetPropertyValue("IJUAhsAngle1", "Angle1").ToString)
                        If d1 < 0.0001 Then
                            oVec = vector2
                        ElseIf d2 < 0.0001 Then
                            oVec = vector3
                        ElseIf angle1 < Math.PI Then
                            oVec = vector2
                        ElseIf angle1 > Math.PI Then
                            oVec = vector3
                        End If
                        For Each oPipeTapFeature In oPipe.GetRelationship("HasTapFeatures", "TapFeature").TargetObjects
                            Dim oRelationColl As RelationCollection = oPipeTapFeature.GetRelationship("GeneratesTap", "GeneratedTap")
                            Dim oDistribPort As IDistributionPort
                            oDistribPort = oRelationColl.TargetObjects.Item(0)
                            run2 = oRelationColl.TargetObjects.Item(0).GetRelationship("FlowPorts"， "DistribConnectionObj").TargetObjects.Item(0).GetRelationship("OwnsDistributionConnection"， "DistribParent").TargetObjects.Item(0)
                            For Each DummyPipe As PipeStraightFeature In run2.SystemChildren.OfType(Of PipeStraightFeature)
                                Dim vector1 As Vector = New Vector(DummyPipe.EndLocation - DummyPipe.StartLocation)
                                If DummyPipe.Parts.Item(0).Name = "Dummy Pipe" And (vector1.Angle(oVec, vector6) < 0.0001 Or vector1.Angle(oVec, vector4 - vector6) < 0.0001) Then
                                    run2.Delete()
                                    ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                    oPipeTapFeature.Delete()
                                    ClientServiceProvider.TransactionMgr.Commit("del dummypipe")
                                End If
                            Next
                        Next
                    Else
                        ToolStripStatusLabel1.Text = "耳轴位置不对"
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                ToolStripStatusLabel1.Text = "选一个管架"
            End If
        Else
            ToolStripStatusLabel1.Text = "选一个管架"
        End If
    End Sub

    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        m_UOM = MiddleServiceProvider.UOMMgr
        If oSelect.SelectedObjects.Count = 1 Then
            If TypeOf (oSelect.SelectedObjects.Item(0)) Is Support Then
                Dim oSupport As Support = oSelect.SelectedObjects.Item(0)
                Dim Dummy As SupportComponent
                Dim OPD As PropertyValueDouble = Nothing
                Dim oLength As PropertyValueDouble = Nothing
                Dim oLengthd As Double
                Dim TapNum As String = Nothing
                Dim oPipe As RouteFeature = oSupport.SupportedObjects.Item(0)
                Dim run1 As PipeRun = oPipe.SystemParent
                Dim oPipeline As Pipeline = run1.SystemParent
                Dim oTurnEnd As Position = New Position()
                Dim vector1 As Vector = New Vector()
                Dim vector2 As Vector, vector3 As Vector
                Dim vector4 As Vector = New Vector()
                Dim oRefport As RefPortHelper = New RefPortHelper(oSupport)
                Dim oMat As Matrix4X4 = oRefport.PortLCS("Route")
                Dim oVec As Vector = oMat.ZAxis
                Dim oVecX As Vector = oMat.XAxis
                Try
                    If oPipe.ClassInfo.Name = "CPPipeTurnPathFeat" Then
                        Dim I As Integer = 0
                        For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                            If Dummy.BOMDescription.Contains("PIPE") Then
                                I += 1
                            End If
                        Next
                        If I = 0 Then
                            ToolStripStatusLabel1.Text = "不是耳轴"
                            Exit Sub
                        Else
                            Dim Orr(I - 1) As Double
                            Dim Arr(I - 1) As Double
                            Dim J As Integer = 0
                            For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                                If Dummy.BOMDescription.Contains("PIPE") Then
                                    OPD = Dummy.GetPropertyValue("IJOAhsPipeOD", "PipeOD")
                                    oLength = Dummy.GetPropertyValue("IJUAhsHeight1", "Height1")
                                    Arr(J) = m_UOM.ParseUnit(UnitType.Distance, oLength.ToString)
                                    Orr(J) = m_UOM.ParseUnit(UnitType.Distance, OPD.ToString)
                                    J += 1
                                End If
                            Next
                            If Arr.Count > 1 Then
                                For k = 1 To I - 1
                                    If Arr（k) < Arr(0) And Arr(k) <> 0.0001 Then
                                        If Arr(k) <> 0.11 And Orr(k) = Orr(0) Then
                                            Arr(0) = Arr(k)
                                            Orr(0) = Orr(k)
                                        End If
                                    End If
                                Next
                            End If
                            oLengthd = Arr(0)
                            Dim OD As Double
                            OD = Orr(0)
                            Dim NS As Double
                            Dim DIN As Double
                            TapNum = TapOD(OD, NS, DIN)
                            If TapNum = Nothing Then
                                ToolStripStatusLabel1.Text = "找不到尺寸"
                                Exit Sub
                            End If
                            Dim run2 As PipeRun
                            If run1.NPD.Units = "mm" Then
                                run2 = New PipeRun(oPipeline, DIN, "mm", run1.Specification.SpecificationName)
                            Else
                                run2 = New PipeRun(oPipeline, NS, "in", run1.Specification.SpecificationName)
                            End If
                            run2.FlowDirection = 4
                            vector2 = oPipe.Location - oPipe.StartLocation
                            vector3 = oPipe.Location - oPipe.EndLocation
                            Dim d1 As Double = Math.Abs(oSupport.SupportPosition.DistanceToPoint(oPipe.StartLocation))
                            Dim d2 As Double = Math.Abs(oSupport.SupportPosition.DistanceToPoint(oPipe.EndLocation))
                            Dim angle1 As Double = 0
                            angle1 = m_UOM.ParseUnit(UnitType.Angle, oSupport.GetRelationship("AssemblyHierarchy", "IsChildOf").TargetObjects(0).GetPropertyValue("IJUAhsAngle1", "Angle1").ToString)
                            If d1 < 0.0001 Then
                                oTurnEnd = oPipe.StartLocation
                                oVec = vector2
                            ElseIf d2 < 0.0001 Then
                                oTurnEnd = oPipe.EndLocation
                                oVec = vector3
                            ElseIf angle1 < Math.PI Then
                                oTurnEnd = oPipe.StartLocation
                                oVec = vector2
                            ElseIf angle1 > Math.PI Then
                                oTurnEnd = oPipe.EndLocation
                                oVec = vector3
                            End If
                            Dim oPTapDirection As PipeTapDirection = PipeTapDirection.PipeTapDirection_PARALLEL
                            Dim oPipeTapFeature As PipeTapFeature
                            oPipeTapFeature = New PipeTapFeature(oPipe, TapNum, oTurnEnd, oVec, oPTapDirection)
                            If oTurnEnd = oPipe.EndLocation Then
                                oPipeTapFeature.ReferencePortIndex = 2
                                oPipeTapFeature.DistanceFromReference = 0
                            End If
                            ClientServiceProvider.TransactionMgr.Commit("new tap")
                            Dim oPEF As PipeEndFeature = Nothing
                            Dim oRelationColl As RelationCollection = oPipeTapFeature.GetRelationship("GeneratesTap", "GeneratedTap")
                            Dim oDistribPort As IDistributionPort
                            oDistribPort = oRelationColl.TargetObjects.Item(0)
                            oVec.Length = oLengthd
                            Dim oPosition As Position = New Position（oVec.X + oTurnEnd.X, oVec.Y + oTurnEnd.Y, oVec.Z + oTurnEnd.Z）
                            run2.RouteBetweenPortAndPoint(oDistribPort, oPosition, oPEF)
                            ClientServiceProvider.TransactionMgr.Compute()
                            If My.Settings.note.Replace(" ", "").Trim IsNot String.Empty Then
                                Dim oFeatPosition As Position = Nothing
                                Dim oFeat As StraightFeature = Nothing
                                run2.GetFeatureAtLocation(PathFeatureObjectTypes.PathFeatureType_STRAIGHT, PathFeatureFunctions.PathFeatureFunction_ROUTE, oPEF.Location, oFeatPosition, oFeat)
                                oFeat.Parts.Item(0).SetPropertyValue("Dummy Pipe", "IJNamedItem", "Name")
                                oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingRequirement, "IJMTOInfo", "ReportingRequirements")
                                oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingType, "IJMTOInfo", "ReportingType")
                                Dim oControlPoint As New ControlPoint(oFeat.Parts.Item(0), oFeat.StartLocation, 0.05)
                                oControlPoint.SubType = Int(ControlPointSubType.SupportNotes)
                                ClientServiceProvider.TransactionMgr.Commit("new controlpoint")
                                oControlPoint.SetUserDefinedName("ControlPnt1")
                                Dim oNote As Note
                                oNote = New Note(oControlPoint)
                                oNote.SetPropertyValue(3, "IJGeneralNote", "Purpose")
                                oNote.SetPropertyValue(My.Settings.note, "IJGeneralNote", "Text")
                                oNote.SetPropertyValue("TAPnote", "IJGeneralNote", "Name")
                                oNote.AddRelationToObject(oControlPoint)
                            End If
                            ClientServiceProvider.TransactionMgr.Commit("new dummypipe")
                        End If
                    ElseIf oPipe.ClassInfo.Name = "CPPipeStraightPathFeat" Then
                        If oSupport.Name.Contains("S033") Or oSupport.Name.Contains("S034") Or oSupport.Name.Contains("S035") Or oSupport.Name.Contains("S036") Or oSupport.Name.Contains("S076") Or oSupport.Name.Contains("S077") Or oSupport.Name.Contains("S078") Or oSupport.Name.Contains("S079") Or oSupport.Name.Contains("S063") Or oSupport.Name.Contains("S265") Or oSupport.Name.Contains("S266") Or oSupport.Name.Contains("S267") Then
                            Dim I As Integer = 0
                            For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                                If Dummy.BOMDescription.Contains("PIPE") Then
                                    I += 1
                                End If
                            Next
                            Dim Orr(I - 1) As Double
                            Dim Arr(I - 1) As Double
                            Dim J As Integer = 0
                            For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                                If Dummy.BOMDescription.Contains("PIPE") Then
                                    OPD = Dummy.GetPropertyValue("IJOAhsPipeOD", "PipeOD")
                                    oLength = Dummy.GetPropertyValue("IJUAhsHeight1", "Height1")
                                    Arr(J) = m_UOM.ParseUnit(UnitType.Distance, oLength.ToString)
                                    Orr(J) = m_UOM.ParseUnit(UnitType.Distance, OPD.ToString)
                                    J += 1
                                End If
                            Next
                            If Arr.Count > 1 Then
                                For k = 1 To I - 1
                                    If Arr（k) < Arr(0) And Arr(k) <> 0.0001 Then
                                        If Arr(k) <> 0.11 And Orr(k) = Orr(0) Then
                                            Arr(0) = Arr(k)
                                            Orr(0) = Orr(k)
                                        End If
                                    End If
                                Next
                            End If
                            oLengthd = Arr(0)
                            Dim OD As Double
                            OD = Orr(0)
                            Dim NS As Double
                            Dim DIN As Double
                            TapNum = TapOD(OD, NS, DIN)
                            If TapNum = Nothing Then
                                ToolStripStatusLabel1.Text = "找不到尺寸"
                                Exit Sub
                            End If
                            Dim run2 As PipeRun
                            If run1.NPD.Units = "mm" Then
                                run2 = New PipeRun(oPipeline, DIN, "mm", run1.Specification.SpecificationName)
                            Else
                                run2 = New PipeRun(oPipeline, NS, "in", run1.Specification.SpecificationName)
                            End If
                            run2.FlowDirection = 4
                            Dim oPTapDirection As PipeTapDirection = PipeTapDirection.PipeTapDirection_PERPENDICULAR
                            Dim position1 As Position = oSupport.SupportPosition
                            Dim Position2 As Position = New Position(position1.X + oVec.X, position1.Y + oVec.Y, position1.Z + oVec.Z)
                            Dim oPipeTapFeature As PipeTapFeature = New PipeTapFeature(oPipe, TapNum, position1, vector4, oPTapDirection)
                            ClientServiceProvider.TransactionMgr.Compute()
                            Dim vector5 As Vector
                            vector5 = oPipeTapFeature.Direction
                            Dim angle1 As Double = 0
                            If oSupport.Name.Contains("S063") Then
                                angle1 = m_UOM.ParseUnit(UnitType.Angle, oSupport.GetPropertyValue("IJUAhsAngle1", "Angle1").ToString)
                            End If
                            Dim oMat1 As Matrix4X4 = New Matrix4X4(oMat)
                            Dim vector6 As Vector = oPipe.EndLocation - oPipe.StartLocation
                            If oVecX.Add(vector6).Length > oVecX.Subtract(vector6).Length Then
                                angle1 = -angle1
                            End If
                            oVec = oMat1.ZAxis
                            oVec = MyRotate(vector6, oVec, -angle1)
                            oPipeTapFeature.Angle = -vector5.Angle(oVec, vector6)
                            ClientServiceProvider.TransactionMgr.Commit("new tap")
                            Dim oPEF As PipeEndFeature = Nothing
                            Dim oRelationColl As RelationCollection = oPipeTapFeature.GetRelationship("GeneratesTap", "GeneratedTap")
                            Dim oDistribPort As IDistributionPort
                            oDistribPort = oRelationColl.TargetObjects.Item(0)
                            Dim oVec1 As Vector = oDistribPort.Location - position1
                            oVec1.Length = oLengthd
                            Dim oPosition As Position = New Position（oVec1.X + position1.X, oVec1.Y + position1.Y, oVec1.Z + position1.Z）
                            run2.RouteBetweenPortAndPoint(oDistribPort, oPosition, oPEF)
                            ClientServiceProvider.TransactionMgr.Compute()
                            If My.Settings.note.Replace(" ", "").Trim IsNot String.Empty Then
                                Dim oFeatPosition As Position = Nothing
                                Dim oFeat As StraightFeature = Nothing
                                run2.GetFeatureAtLocation(PathFeatureObjectTypes.PathFeatureType_STRAIGHT, PathFeatureFunctions.PathFeatureFunction_ROUTE, oPEF.Location, oFeatPosition, oFeat)
                                oFeat.Parts.Item(0).SetPropertyValue("Dummy Pipe", "IJNamedItem", "Name")
                                oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingRequirement, "IJMTOInfo", "ReportingRequirements")
                                oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingType, "IJMTOInfo", "ReportingType")
                                Dim oControlPoint As New ControlPoint(oFeat.Parts.Item(0), oFeat.StartLocation, 0.05)
                                oControlPoint.SubType = Int(ControlPointSubType.SupportNotes)
                                ClientServiceProvider.TransactionMgr.Commit("new controlpoint")
                                oControlPoint.SetUserDefinedName("ControlPnt1")
                                Dim oNote As Note
                                oNote = New Note(oControlPoint)
                                oNote.SetPropertyValue(3, "IJGeneralNote", "Purpose")
                                oNote.SetPropertyValue(My.Settings.note, "IJGeneralNote", "Text")
                                oNote.SetPropertyValue("TAPnote", "IJGeneralNote", "Name")
                                oNote.AddRelationToObject(oControlPoint)
                            End If
                            ClientServiceProvider.TransactionMgr.Commit("new dummypipe")
                        Else
                            Dim Sum As Integer = 0
                            For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                                If Dummy.BOMDescription.Contains("PIPE") Then
                                    Sum += 1
                                End If
                            Next
                            If Sum = 0 Then
                                ToolStripStatusLabel1.Text = "不是耳轴"
                                Exit Sub
                            Else
                                Dim angle As Double = 2 * Math.PI / Sum
                                Dim I As Integer = 0
                                For Each Dummy In oSupport.SystemChildren.OfType(Of SupportComponent)
                                    If Dummy.BOMDescription.Contains("PIPE") Then
                                        OPD = Dummy.GetPropertyValue("IJOAhsPipeOD", "PipeOD")
                                        oLength = Dummy.GetPropertyValue("IJUAhsHeight1", "Height1")
                                        Dim OD As Double
                                        OD = m_UOM.ParseUnit(UnitType.Distance, OPD.ToString)
                                        Dim NS As Double
                                        Dim DIN As Double
                                        TapNum = TapOD(OD, NS, DIN)
                                        If TapNum = Nothing Then
                                            ToolStripStatusLabel1.Text = "找不到尺寸"
                                            Exit Sub
                                        End If
                                        Dim run2 As PipeRun
                                        If run1.NPD.Units = "mm" Then
                                            run2 = New PipeRun(oPipeline, DIN, "mm", run1.Specification.SpecificationName)
                                        Else
                                            run2 = New PipeRun(oPipeline, NS, "in", run1.Specification.SpecificationName)
                                        End If
                                        run2.FlowDirection = 4
                                        Dim oPTapDirection As PipeTapDirection = PipeTapDirection.PipeTapDirection_PERPENDICULAR
                                        Dim position1 As Position = New Position(oSupport.SupportPosition)
                                        Dim vector6 As Vector = oPipe.EndLocation - oPipe.StartLocation
                                        Dim angle1 As Double = 0
                                        If oSupport.Name.Contains("S031") Or oSupport.Name.Contains("S264") Then
                                            angle1 = m_UOM.ParseUnit(UnitType.Angle, oSupport.GetPropertyValue("IJUAhsAngle1", "Angle1").ToString)
                                        End If
                                        If oSupport.Name.Contains("S030") Or oSupport.Name.Contains("S031") Or oSupport.Name.Contains("S264") Then
                                            Dim Offset1 As Double
                                            Offset1 = m_UOM.ParseUnit(UnitType.Distance, oSupport.GetPropertyValue("IJUAhsOffset1", "Offset1").ToString)
                                            vector6.Length = OD / 2 + Offset1
                                            If oVecX.Add(vector6).Length < oVecX.Subtract(vector6).Length Then
                                                vector6.Length = -vector6.Length
                                            Else
                                                angle1 = -angle1
                                            End If
                                            position1.Set(vector6.X + position1.X, vector6.Y + position1.Y, vector6.Z + position1.Z)
                                            Dim oPipePosition As Position = Nothing
                                            run1.GetFeatureAtLocation(PathFeatureObjectTypes.PathFeatureType_STRAIGHT, PathFeatureFunctions.PathFeatureFunction_ROUTE, position1, oPipePosition, oPipe)
                                        Else
                                            position1 = oSupport.SupportPosition
                                        End If
                                        Dim oMat1 As Matrix4X4 = New Matrix4X4(oMat)
                                        oMat1.Rotate(I * angle, vector6)
                                        oVec = oMat1.ZAxis
                                        oVec = MyRotate(vector6, oVec, -angle1)
                                        Dim oPipeTapFeature As PipeTapFeature = New PipeTapFeature(oPipe, TapNum, position1, vector4, oPTapDirection)
                                        ClientServiceProvider.TransactionMgr.Compute()
                                        Dim vector5 As Vector
                                        vector5 = oPipeTapFeature.Direction
                                        oPipeTapFeature.Angle = -vector5.Angle(oVec, vector6)
                                        ClientServiceProvider.TransactionMgr.Commit("new tap")
                                        Dim oPEF As PipeEndFeature = Nothing
                                        Dim oRelationColl As RelationCollection = oPipeTapFeature.GetRelationship("GeneratesTap", "GeneratedTap")
                                        Dim oDistribPort As IDistributionPort
                                        oDistribPort = oRelationColl.TargetObjects.Item(0)
                                        Dim oVec1 As Vector = oDistribPort.Location - position1
                                        oVec1.Length = m_UOM.ParseUnit(UnitType.Distance, oLength.ToString)
                                        Dim oPosition As Position = New Position（oVec1.X + position1.X, oVec1.Y + position1.Y, oVec1.Z + position1.Z）
                                        run2.RouteBetweenPortAndPoint(oDistribPort, oPosition, oPEF)
                                        ClientServiceProvider.TransactionMgr.Compute()
                                        If My.Settings.note.Replace(" ", "").Trim IsNot String.Empty Then
                                            Dim oFeatPosition As Position = Nothing
                                            Dim oFeat As StraightFeature = Nothing
                                            run2.GetFeatureAtLocation(PathFeatureObjectTypes.PathFeatureType_STRAIGHT, PathFeatureFunctions.PathFeatureFunction_ROUTE, oPEF.Location, oFeatPosition, oFeat)
                                            oFeat.Parts.Item(0).SetPropertyValue("Dummy Pipe", "IJNamedItem", "Name")
                                            oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingRequirement, "IJMTOInfo", "ReportingRequirements")
                                            oFeat.Parts.Item(0).SetPropertyValue(My.Settings.ReportingType, "IJMTOInfo", "ReportingType")
                                            Dim oControlPoint As New ControlPoint(oFeat.Parts.Item(0), oFeat.StartLocation, 0.05)
                                            oControlPoint.SubType = Int(ControlPointSubType.SupportNotes)
                                            ClientServiceProvider.TransactionMgr.Commit("new controlpoint")
                                            oControlPoint.SetUserDefinedName("ControlPnt" + CStr(I + 1))
                                            Dim oNote As Note
                                            oNote = New Note(oControlPoint)
                                            oNote.SetPropertyValue(3, "IJGeneralNote", "Purpose")
                                            oNote.SetPropertyValue(My.Settings.note, "IJGeneralNote", "Text")
                                            oNote.SetPropertyValue("TAPnote", "IJGeneralNote", "Name")
                                            oNote.AddRelationToObject(oControlPoint)
                                        End If
                                        ClientServiceProvider.TransactionMgr.Commit("new dummypipe")
                                        I += 1
                                    End If
                                Next
                            End If
                        End If
                    Else
                        ToolStripStatusLabel1.Text = "耳轴位置不对"
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                ToolStripStatusLabel1.Text = "选一个管架"
            End If
        Else
            ToolStripStatusLabel1.Text = "选一个管架"
        End If
    End Sub
End Class