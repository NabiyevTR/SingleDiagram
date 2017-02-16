

Imports System.Windows.Forms

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Imports Microsoft.Office.Interop




Public Class ADM_project_1


    <CommandMethod("dSLD")>
    Public Shared Sub dSLD()



        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor
        Dim myform As System.Windows.Forms.Form
        myform = New Form1
        myform.Show()



    End Sub

















 

End Class