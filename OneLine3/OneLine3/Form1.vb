

Imports Microsoft.Office.Interop
Imports System.Math
Imports System.Windows.Forms

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput


Public Class Form1

    Structure FeaderData

        Public Number As String

        Public Label As String

        Public PlanLabel As String

        Public Name As String

        Public Destination As String

        Public Phase As String

        Public Voltage As String

        Public InstalledActivePower As String

        Public PowerFactor As String

        Public MaxFactor As String

        Public DemandFactor As String

        Public RatedActivePower As String

        Public RatedReactivePower As String

        Public RatedFullPower As String

        Public RatedCurrent As String

        Public SwitchType As String

        Public SwitchNominalCurrent As String

        Public SwitchRelease As String

        Public SwitchReleaseCurrent As String

        Public RCDSetPoint As String

        Public RCDType As String

        Public Contactor As String

        Public ContactorType As String

        Public Cable As String

        Public EstimatedLength As String

        Public CableVoltageDeviation As String

        Public FullVoltageDeviation As String

        Public ThreePhShortCurrentMax As String

        Public ThreePhShortCurrentMin As String

        Public SinglePhShortCurrentMax As String

        Public SinglePhShortCurrentMin As String

    End Structure


    Structure db

        Public Number As String

        Public Label As String

        Public PlanLabel as String

        Public dbName As String

        Public Name As String

        Public Destination As String

        Public Phase As String

        Public Voltage As String

        Public InstalledActivePowerA As String
        Public InstalledActivePowerB As String
        Public InstalledActivePowerC As String
        Public InstalledActivePowerMax As String
        Public InstalledActivePower As String

        Public PowerFactor As String

        Public MaxFactor As String

        Public RatedActivePowerA As String
        Public RatedActivePowerB As String
        Public RatedActivePowerC As String
        Public RatedActivePowerMax As String
        Public RatedActivePower As String

        Public RatedReactivePowerA As String
        Public RatedReactivePowerB As String
        Public RatedReactivePowerC As String
        Public RatedReactivePowerMax As String
        Public RatedReactivePower As String

        Public RatedFullPowerA As String
        Public RatedFullPowerB As String
        Public RatedFullPowerC As String
        Public RatedFullPowerMax As String
        Public RatedFullPower As String

        Public RatedCurrentA As String
        Public RatedCurrentB As String
        Public RatedCurrentC As String
        Public RatedCurrentMax As String
        Public RatedCurrent As String

        Public InstalledPowerNonSym As String
        Public RatedCurrentNonSym As String

        Public SwitchType As String

        Public SwitchNominalCurrent As String

        Public SwitchRelease As String

        Public SwitchReleaseCurrent As String

        Public RCDSetPoint As String

        Public RCDType As String

        Public ContactorType As String

        Public Cable As String

        Public EstimatedLength As String

        Public CableVoltageDeviation As String

        Public FullVoltageDeviation As String

        Public ThreePhShortCurrentMax As String

        Public ThreePhShortCurrentMin As String

        Public SinglePhShortCurrentMax As String

        Public SinglePhShortCurrentMin As String

        Public StartF As Integer

        Public StopF As Integer

        Public FeadersCount As Integer

        Public Feaders() As FeaderData

    End Structure





    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OFD1.FileOk

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGetFileName.Click


        ' Указываем начальную папку
        OFD1.InitialDirectory = "C:"
        ' Указываем заголовок
        OFD1.Title = "Файл с расчетами"
        'Фильтр Файлов
        OFD1.Filter = "Excel файлы|*.xls; *.xlsx; *.xlsm"

        If OFD1.ShowDialog = DialogResult.OK Then DataFileName.Text = OFD1.FileName Else Exit Sub


        'получаем доступ к Excel
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook


        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(DataFileName.Text)
        'xlWorkSheet = xlWorkBook.Worksheets("sheet1")


        'записываем в lsbxExcel
        For Each sheet As Excel.Worksheet In xlApp.ActiveWorkbook.Worksheets
            lsbxExcel.Items.Add(sheet.Name)
        Next







        ' закрываем Excel
        xlWorkBook.Close(False)
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)




    End Sub


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub btClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClose.Click
        Me.Close()
    End Sub

    Private Sub btIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btIn.Click


        'Выбор листов для импорта в AutoCAD
        For Each itm As Object In lsbxExcel.SelectedItems
            lsbxAutocad.Items.Add(itm.ToString)
        Next
        'Снимаем выделение с lsbxExcel
        lsbxExcel.SelectedItem = Nothing




    End Sub

    Private Sub lsbxExcel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lsbxExcel.SelectedIndexChanged

    End Sub

    Private Sub btOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btOut.Click


        For i As Integer = lsbxAutocad.SelectedIndices.Count - 1 To 0 Step -1
            lsbxAutocad.Items.RemoveAt(lsbxAutocad.SelectedIndices.Item(i))
        Next


    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGenerate.Click
        'получаем доступ к Excel
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(DataFileName.Text)


        Dim lsbxitem As String
        'определение количества щитов
        Dim DBcount As Integer = lsbxAutocad.Items.Count
        'задаем массив щитов
        Dim MyDB(DBcount - 1) As DB
        'объявлем счетчик
        Dim icount As Integer = 0

        'цикл перебора выбранных щитов
        For Each lsbxitem In lsbxAutocad.Items
            xlWorkSheet = xlWorkBook.Worksheets(lsbxitem.ToString)


            'считываем параметры из xl
            'ВНИМАНИЕ!!! Считывание параметров жестко завязано на ячейки таблицы. 


            With MyDB(icount)

                'название щита
                Try
                    If lsbxitem <> Nothing Then .dbName = lsbxitem.ToString
                Catch ex As Exception
                End Try

                'маркировка
                Try
                    If xlWorkSheet.Range("E9").Value <> Nothing Then .Label = xlWorkSheet.Range("E9").Value.ToString
                Catch ex As Exception
                End Try

                'Наименование подключаемой нагрузки
                Try
                    If xlWorkSheet.Range("D9").Value <> Nothing Then .Name = xlWorkSheet.Range("D9").Value.ToString
                Catch ex As Exception
                End Try

                'Номер автомата
                Try
                    If xlWorkSheet.Range("A9").Value <> Nothing Then .Name = xlWorkSheet.Range("A9").Value.ToString
                Catch ex As Exception
                End Try

                Try
                    If xlWorkSheet.Range("B9").Value <> Nothing Then .Name = .Name + xlWorkSheet.Range("B9").Value.ToString
                Catch ex As Exception
                End Try

                'Расположение
                Try
                    If xlWorkSheet.Range("E9").Value <> Nothing Then .Destination = xlWorkSheet.Range("E9").Value.ToString
                Catch ex As Exception
                End Try

                'Фазность
                Try
                    If xlWorkSheet.Range("AD9").Value <> Nothing Then .Phase = xlWorkSheet.Range("AD9").Value.ToString
                Catch ex As Exception
                End Try

                'напряжение
                Try
                    If xlWorkSheet.Range("AE9").Value <> Nothing Then .Voltage = xlWorkSheet.Range("AE9").Value.ToString
                Catch ex As Exception
                End Try


                'Установленная активная мощность в фазе A
                Try
                    If xlWorkSheet.Range("M9").Value <> Nothing Then .InstalledActivePowerA = MyMod.RoundSign(xlWorkSheet.Range("M9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Установленная активная мощность в фазе B
                Try
                    If xlWorkSheet.Range("M10").Value <> Nothing Then .InstalledActivePowerB = MyMod.RoundSign(xlWorkSheet.Range("M10").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Установленная активная мощность в фазе C
                Try
                    If xlWorkSheet.Range("M11").Value <> Nothing Then .InstalledActivePowerC = MyMod.RoundSign(xlWorkSheet.Range("M11").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Установленная активная мощность максимальная
                Try
                    If xlWorkSheet.Range("AF9").Value <> Nothing Then .InstalledActivePowerMax = MyMod.RoundSign(xlWorkSheet.Range("AF9").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Установленная активная мощность суммарная
                Try
                    If xlWorkSheet.Range("M9").Value <> Nothing And xlWorkSheet.Range("M10").Value <> Nothing And xlWorkSheet.Range("M11").Value <> Nothing Then .InstalledActivePower = MyMod.RoundSign(xlWorkSheet.Range("M9").Value + xlWorkSheet.Range("M10").Value + xlWorkSheet.Range("M11").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Косинус 
                Try
                    If xlWorkSheet.Range("AG9").Value <> Nothing Then .PowerFactor = MyMod.RoundSign(xlWorkSheet.Range("AG9").Value, 2).ToString
                Catch ex As Exception
                End Try

                'Коэффициент одновременности 
                Try
                    If xlWorkSheet.Range("U9").Value <> Nothing And xlWorkSheet.Range("U10").Value <> Nothing And xlWorkSheet.Range("U11").Value <> Nothing And xlWorkSheet.Range("AF9").Value <> Nothing Then .MaxFactor = MyMod.RoundSign((xlWorkSheet.Range("U9").Value + xlWorkSheet.Range("U10").Value + xlWorkSheet.Range("U11").Value) / xlWorkSheet.Range("AF9").Value, 2).ToString
                Catch ex As Exception
                End Try


                'Расчетная активная мощность в фазе A
                Try
                    If xlWorkSheet.Range("U9").Value <> Nothing Then .RatedActivePowerA = MyMod.RoundSign(xlWorkSheet.Range("U9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная активная мощность в фазе B
                Try
                    If xlWorkSheet.Range("U10").Value <> Nothing Then .RatedActivePowerB = MyMod.RoundSign(xlWorkSheet.Range("U10").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная активная мощность в фазе C
                Try
                    If xlWorkSheet.Range("U11").Value <> Nothing Then .RatedActivePowerC = MyMod.RoundSign(xlWorkSheet.Range("U11").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная активная мощность максимальная
                Try
                    If xlWorkSheet.Range("AH9").Value <> Nothing Then .RatedActivePowerMax = MyMod.RoundSign(xlWorkSheet.Range("AH9").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная активная мощность суммарная
                Try
                    If xlWorkSheet.Range("U9").Value <> Nothing And xlWorkSheet.Range("U10").Value <> Nothing And xlWorkSheet.Range("U11").Value <> Nothing Then .RatedActivePower = MyMod.RoundSign(xlWorkSheet.Range("U9").Value + xlWorkSheet.Range("U10").Value + xlWorkSheet.Range("U11").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная реактивная мощность в фазе A
                Try
                    If xlWorkSheet.Range("V9").Value <> Nothing Then .RatedReactivePowerA = MyMod.RoundSign(xlWorkSheet.Range("V9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная реактивная мощность в фазе B
                Try
                    If xlWorkSheet.Range("V10").Value <> Nothing Then .RatedReactivePowerB = MyMod.RoundSign(xlWorkSheet.Range("V10").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная реактивная мощность в фазе C
                Try
                    If xlWorkSheet.Range("V11").Value <> Nothing Then .RatedReactivePowerC = MyMod.RoundSign(xlWorkSheet.Range("V11").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная реактивная мощность максимальная
                Try
                    If xlWorkSheet.Range("AI9").Value <> Nothing Then .RatedReactivePowerMax = MyMod.RoundSign(xlWorkSheet.Range("AI9").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная реактивная мощность суммарная
                Try
                    If xlWorkSheet.Range("V9").Value <> Nothing And xlWorkSheet.Range("V10").Value <> Nothing And xlWorkSheet.Range("V11").Value <> Nothing Then .RatedReactivePower = MyMod.RoundSign(xlWorkSheet.Range("V9").Value + xlWorkSheet.Range("V10").Value + xlWorkSheet.Range("V11").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная полная мощность в фазе A
                Try
                    If xlWorkSheet.Range("W9").Value <> Nothing Then .RatedFullPowerA = MyMod.RoundSign(xlWorkSheet.Range("W9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная полная мощность в фазе B
                Try
                    If xlWorkSheet.Range("W10").Value <> Nothing Then .RatedFullPowerB = MyMod.RoundSign(xlWorkSheet.Range("W10").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная полная мощность в фазе C
                Try
                    If xlWorkSheet.Range("W11").Value <> Nothing Then .RatedFullPowerC = MyMod.RoundSign(xlWorkSheet.Range("W11").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетная полная мощность максимальная
                Try
                    If xlWorkSheet.Range("AJ9").Value <> Nothing Then .RatedFullPowerMax = MyMod.RoundSign(xlWorkSheet.Range("AJ9").Value, 1).ToString
                Catch ex As Exception
                End Try

                'Расчетная полная мощность суммарная
                Try
                    If xlWorkSheet.Range("U9").Value <> Nothing And xlWorkSheet.Range("U10").Value <> Nothing And xlWorkSheet.Range("U11").Value <> Nothing Then .RatedFullPower = MyMod.RoundSign(Sqrt((xlWorkSheet.Range("U9").Value + xlWorkSheet.Range("U10").Value + xlWorkSheet.Range("U11").Value) ^ 2 + (xlWorkSheet.Range("V9").Value + xlWorkSheet.Range("V10").Value + xlWorkSheet.Range("V11").Value) ^ 2), 1).ToString
                Catch ex As Exception
                End Try


                'Расчетный ток в фазе A
                Try
                    If xlWorkSheet.Range("X9").Value <> Nothing Then .RatedCurrentA = MyMod.RoundSign(xlWorkSheet.Range("X9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетный ток в фазе B
                Try
                    If xlWorkSheet.Range("X10").Value <> Nothing Then .RatedCurrentB = MyMod.RoundSign(xlWorkSheet.Range("X10").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетный ток в фазе C
                Try
                    If xlWorkSheet.Range("X11").Value <> Nothing Then .RatedCurrentC = MyMod.RoundSign(xlWorkSheet.Range("X11").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетный ток максимальный
                Try
                    If xlWorkSheet.Range("AK9").Value <> Nothing Then .RatedCurrentMax = MyMod.RoundSign(xlWorkSheet.Range("AK9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Расчетный ток суммарный
                Try
                    .RatedCurrent = MyMod.RoundSign(Sqrt(((xlWorkSheet.Range("U9").Value + xlWorkSheet.Range("U10").Value + xlWorkSheet.Range("U11").Value) ^ 2 + (xlWorkSheet.Range("V9").Value + xlWorkSheet.Range("V10").Value + xlWorkSheet.Range("V11").Value) ^ 2)) / (Sqrt(3) * 0.38), 1).ToString
                Catch ex As Exception
                End Try


                'Нессиметрия по установленной мощности
                Try
                    If xlWorkSheet.Range("AF12").Value <> Nothing Then .InstalledPowerNonSym = MyMod.RoundSign(xlWorkSheet.Range("AF12").Value, 0).ToString
                Catch ex As Exception
                End Try


                'Несимметрияя по расчетному току
                Try
                    If xlWorkSheet.Range("AJ12").Value <> Nothing Then .RatedCurrentNonSym = MyMod.RoundSign(xlWorkSheet.Range("AJ12").Value, 0).ToString
                Catch ex As Exception
                End Try


                'Тип выключателя
                Try
                    If xlWorkSheet.Range("AZ9").Value <> Nothing Then .SwitchType = xlWorkSheet.Range("AZ9").Value.ToString
                Catch ex As Exception
                End Try


                'Номинальный ток выключателя
                Try
                    If xlWorkSheet.Range("BA9").Value <> Nothing Then .SwitchNominalCurrent = xlWorkSheet.Range("BA9").Value.ToString
                Catch ex As Exception
                End Try


                'Тип расцепителя
                Try
                    If xlWorkSheet.Range("BB9").Value <> Nothing Then .SwitchRelease = xlWorkSheet.Range("BB9").Value.ToString
                Catch ex As Exception
                End Try


                'Ток уставки расцепителя
                Try
                    If xlWorkSheet.Range("BC9").Value <> Nothing Then .SwitchReleaseCurrent = MyMod.RoundSign(xlWorkSheet.Range("BC9").Value, 1).ToString
                Catch ex As Exception
                End Try


                'Ток утечки
                Try
                    If xlWorkSheet.Range("BD9").Value <> Nothing Then .RCDSetPoint = xlWorkSheet.Range("BD9").Value.ToString
                Catch ex As Exception
                End Try

                If .RCDSetPoint <> "" And .RCDSetPoint <> "-" Then .RCDType = "AC"


                'Тип контактора
                Try
                    If xlWorkSheet.Range("BJ9").Value <> Nothing Then .ContactorType = xlWorkSheet.Range("BJ9").Value.ToString
                Catch ex As Exception
                End Try


                'Кабель
                Try
                    If xlWorkSheet.Range("DS9").Value <> Nothing Then .Cable = xlWorkSheet.Range("DS9").Value.ToString
                Catch ex As Exception
                End Try


                'Расчетная длина кабельной линии
                Try
                    If xlWorkSheet.Range("DU9").Value <> Nothing Then .EstimatedLength = MyMod.RoundSign(xlWorkSheet.Range("DU9").Value, 0).ToString
                Catch ex As Exception
                End Try


                'Потери напряжения на кабеле
                Try
                    If xlWorkSheet.Range("EJ9").Value <> Nothing Then .CableVoltageDeviation = MyMod.RoundSign(Math.Abs(xlWorkSheet.Range("EJ9").Value * 100), 2).ToString
                Catch ex As Exception
                End Try


                'Полные потери напряжения в конце линии
                Try
                    If xlWorkSheet.Range("EM9").Value <> Nothing Then .FullVoltageDeviation = MyMod.RoundSign(xlWorkSheet.Range("EM9").Value * 100, 2).ToString
                Catch ex As Exception
                End Try


                'Трехфазный максимальный ток КЗ
                Try
                    If xlWorkSheet.Range("EW9").Value <> Nothing Then .ThreePhShortCurrentMax = MyMod.RoundSign(xlWorkSheet.Range("EW9").Value, 2).ToString
                Catch ex As Exception
                End Try


                'Трехфазный минимальный ток КЗ
                Try
                    If xlWorkSheet.Range("EX9").Value <> Nothing Then .ThreePhShortCurrentMin = MyMod.RoundSign(xlWorkSheet.Range("EX9").Value, 2).ToString
                Catch ex As Exception
                End Try


                'Однофазный максимальный ток КЗ
                Try
                    If xlWorkSheet.Range("FM9").Value <> Nothing Then .SinglePhShortCurrentMax = MyMod.RoundSign(xlWorkSheet.Range("FM9").Value, 2).ToString
                Catch ex As Exception
                End Try


                'Однофазный минимальный ток КЗ
                Try
                    If xlWorkSheet.Range("FM9").Value <> Nothing Then .SinglePhShortCurrentMin = MyMod.RoundSign(xlWorkSheet.Range("FM9").Value, 2).ToString
                Catch ex As Exception
                End Try





            End With


            'определяем первую строку диапазона суммирования


            Dim jCount As Integer = 2
            MyDB(icount).StartF = -1
            MyDB(icount).StopF = -1


            Do
                jCount = jCount + 1

                If xlWorkSheet.Cells(jCount, 4).value = "первая строка диапазона суммирования групповой нагрузки" Then
                    MyDB(icount).StartF = jCount

                End If



            Loop Until xlWorkSheet.Cells(jCount, 4).value = "первая строка диапазона суммирования групповой нагрузки" Or (jCount > 100)


            'определяем последнюю строку диапазона суммирования

            jCount = jCount - 1

            Do
                jCount = jCount + 1
                If xlWorkSheet.Cells(jCount, 4).value = "последняя строка диапазона суммирования групповой нагрузки" Then
                    MyDB(icount).StopF = jCount
                End If
            Loop Until xlWorkSheet.Cells(jCount, 4).value = "последняя строка диапазона суммирования групповой нагрузки" Or (jCount > 1000)




            'Определяем размер
            ReDim MyDB(icount).Feaders(MyDB(icount).StopF - MyDB(icount).StartF)

            Dim fpos As Integer = 0

            For jCount = MyDB(icount).StartF + 1 To MyDB(icount).StopF - 1

                MyDB(icount).FeadersCount = fpos


                With MyDB(icount).Feaders(fpos)

                    ' номер
                    Try
                        If xlWorkSheet.Range("A" & jCount.ToString).Value <> Nothing Then .Number = xlWorkSheet.Range("A" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    Try
                        If xlWorkSheet.Range("B" & jCount.ToString).Value <> Nothing Then .Number = .Number & xlWorkSheet.Range("B" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'маркировка
                    Try
                        If xlWorkSheet.Range("B" & jCount.ToString).Value.ToString <> Nothing Then .Label = xlWorkSheet.Range("B" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    Try
                        If MyDB(icount).dbName <> Nothing Then .Label = .Label & "." & MyDB(icount).dbName
                    Catch ex As Exception
                    End Try

                    'Маркировка по плану
                    Try
                        If xlWorkSheet.Range("C" & jCount.ToString).Value <> Nothing Then .PlanLabel = xlWorkSheet.Range("C" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try


                    'Наименование подключаемой нагрузки
                    Try
                        If xlWorkSheet.Range("D" & jCount.ToString).Value <> Nothing Then .Name = xlWorkSheet.Range("D" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    Try
                        If xlWorkSheet.Range("E" & jCount.ToString).Value <> Nothing Then .Name = .Name + vbCrLf & xlWorkSheet.Range("E" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Расположение
                    Try
                        If xlWorkSheet.Range("E" & jCount.ToString).Value <> Nothing Then .Destination = xlWorkSheet.Range("E" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Установленная мощность
                    Try
                        If xlWorkSheet.Range("AF" & jCount.ToString).Value <> Nothing Then .InstalledActivePower = MyMod.RoundSign(xlWorkSheet.Range("AF" & jCount.ToString).Value, 1).ToString
                    Catch ex As Exception
                    End Try


                    'Определение трехфахной нагрузки в трех строках для определения расчетного тока и расчетной мощности

                    Dim mCells As Integer

                    If xlWorkSheet.Cells(jCount, 4).mergecells Then
                        mCells = xlWorkSheet.Cells(jCount, 4).MergeArea.Cells.Count
                    End If




                    'Расчетный ток
                    Try

                        Dim RatedCurrent As Double = 0


                        If mCells = 3 And xlWorkSheet.Range("F" & jCount.ToString).Value.ToString = "380" Then
                            If xlWorkSheet.Range("X" & jCount.ToString).Value <> Nothing Then RatedCurrent = RatedCurrent + xlWorkSheet.Range("X" & jCount.ToString).Value
                            If xlWorkSheet.Range("X" & (jCount + 1).ToString).Value <> Nothing Then RatedCurrent = RatedCurrent + xlWorkSheet.Range("X" & (jCount + 1).ToString).Value
                            If xlWorkSheet.Range("X" & (jCount + 2).ToString).Value <> Nothing Then RatedCurrent = RatedCurrent + xlWorkSheet.Range("X" & (jCount + 2).ToString).Value

                            .RatedCurrent = MyMod.RoundSign((RatedCurrent / 3), 1).ToString

                        Else
                            If xlWorkSheet.Range("AK" & jCount.ToString).Value <> Nothing Then .RatedCurrent = MyMod.RoundSign(xlWorkSheet.Range("AK" & jCount.ToString).Value, 1).ToString
                        End If


                    Catch ex As Exception
                    End Try



                    

                    'Расчетная мощность

                    Try

                        Dim RatedActivePower As Double = 0


                        If mCells = 3 And xlWorkSheet.Range("F" & jCount.ToString).Value.ToString = "380" Then
                            If xlWorkSheet.Range("U" & jCount.ToString).Value <> Nothing Then RatedActivePower = RatedActivePower + xlWorkSheet.Range("U" & jCount.ToString).Value
                            If xlWorkSheet.Range("U" & (jCount + 1).ToString).Value <> Nothing Then RatedActivePower = RatedActivePower + xlWorkSheet.Range("U" & (jCount + 1).ToString).Value
                            If xlWorkSheet.Range("U" & (jCount + 2).ToString).Value <> Nothing Then RatedActivePower = RatedActivePower + xlWorkSheet.Range("U" & (jCount + 2).ToString).Value

                            .RatedActivePower = MyMod.RoundSign(RatedActivePower, 1).ToString

                        Else
                            If xlWorkSheet.Range("AH" & jCount.ToString).Value <> Nothing Then .RatedActivePower = MyMod.RoundSign(xlWorkSheet.Range("AH" & jCount.ToString).Value, 1).ToString
                        End If








                    Catch ex As Exception
                    End Try





                    '                    Try
                    ' If xlWorkSheet.Range("U" & jCount.ToString).Value <> Nothing Then

                    '  Dim RatedActivePower As Double = 0

                    '  If Not xlWorkSheet.Range("U" & jCount.ToString).MergeCells Then

                    ' For kcount As Integer = jCount To jCount + mCells - 1
                    '   RatedActivePower = RatedActivePower + xlWorkSheet.Range("U" & kcount.ToString).Value
                    '   Next

                    ' Else
                    '    RatedActivePower = xlWorkSheet.Range("U" & jCount.ToString).Value

                    'End If
                    '  .RatedActivePower = MyMod.RoundSign(RatedActivePower, 1).ToString
                    ' End If

                    'Catch ex As Exception
                    ' End Try

                    'Фазность
                    Try
                        If xlWorkSheet.Range("AD" & jCount.ToString).Value <> Nothing Then .Phase = xlWorkSheet.Range("AD" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'напряжение
                    Try
                        If xlWorkSheet.Range("AE" & jCount.ToString).Value <> Nothing Then .Voltage = xlWorkSheet.Range("AE" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try


                    'Косинус 
                    Try
                        If xlWorkSheet.Range("AG" & jCount.ToString).Value <> Nothing Then .PowerFactor = MyMod.RoundSign(xlWorkSheet.Range("AG" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Коэффициент одновременности 
                    Try
                        If xlWorkSheet.Range("AO" & jCount.ToString).Value <> Nothing Then .MaxFactor = MyMod.RoundSign(xlWorkSheet.Range("O" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Коффициент спроса
                    Try
                        If xlWorkSheet.Range("AP" & jCount.ToString).Value <> Nothing Then .DemandFactor = MyMod.RoundSign(xlWorkSheet.Range("P" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Тип выключателя
                    Try
                        If xlWorkSheet.Range("AZ" & jCount.ToString).Value <> Nothing Then .SwitchType = xlWorkSheet.Range("AZ" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Номинальный ток выключателя
                    Try
                        If xlWorkSheet.Range("BA" & jCount.ToString).Value <> Nothing Then .SwitchNominalCurrent = xlWorkSheet.Range("BA" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Тип расцепителя
                    Try
                        If xlWorkSheet.Range("BB" & jCount.ToString).Value <> Nothing Then .SwitchRelease = xlWorkSheet.Range("BB" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Ток уставки расцепителя
                    Try
                        If xlWorkSheet.Range("BC" & jCount.ToString).Value <> Nothing Then .SwitchReleaseCurrent = MyMod.RoundSign(xlWorkSheet.Range("BC" & jCount.ToString).Value, 1).ToString
                    Catch ex As Exception
                    End Try

                    'Ток утечки
                    Try
                        If xlWorkSheet.Range("BD" & jCount.ToString).Value <> Nothing Then .RCDSetPoint = xlWorkSheet.Range("BD" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try
                    If .RCDSetPoint <> "" And .RCDSetPoint <> "-" Then .RCDType = "AC"

                    'маркировка контактора
                    Try
                        If xlWorkSheet.Range("B" & jCount.ToString).Value <> Nothing Then .Contactor = "KM" + xlWorkSheet.Range("B" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Тип контактора
                    Try
                        If xlWorkSheet.Range("BJ" & jCount.ToString).Value <> Nothing Then .ContactorType = xlWorkSheet.Range("BJ" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Кабель
                    Try
                        If xlWorkSheet.Range("DS" & jCount.ToString).Value <> Nothing Then .Cable = xlWorkSheet.Range("DS" & jCount.ToString).Value.ToString
                    Catch ex As Exception
                    End Try

                    'Расчетная длина кабельной линии
                    Try
                        If xlWorkSheet.Range("DU" & jCount.ToString).Value <> Nothing Then .EstimatedLength = MyMod.RoundSign(xlWorkSheet.Range("DU" & jCount.ToString).Value, 0).ToString
                    Catch ex As Exception
                    End Try

                    'Потери напряжения на кабеле
                    Try
                        If xlWorkSheet.Range("EJ" & jCount.ToString).Value <> Nothing Then .CableVoltageDeviation = MyMod.RoundSign(Math.Abs(xlWorkSheet.Range("EJ" & jCount.ToString).Value * 100), 2).ToString
                    Catch ex As Exception
                    End Try

                    'Полные потери напряжения в конце линии
                    Try
                        If xlWorkSheet.Range("EM" & jCount.ToString).Value <> Nothing Then .FullVoltageDeviation = MyMod.RoundSign(xlWorkSheet.Range("EM" & jCount.ToString).Value * 100, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Трехфазный максимальный ток КЗ
                    Try
                        If xlWorkSheet.Range("EW" & jCount.ToString).Value <> Nothing Then .ThreePhShortCurrentMax = MyMod.RoundSign(xlWorkSheet.Range("EW" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Трехфазный минимальный ток КЗ
                    Try
                        If xlWorkSheet.Range("EX" & jCount.ToString).Value <> Nothing Then .ThreePhShortCurrentMin = MyMod.RoundSign(xlWorkSheet.Range("EX" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Однофазный максимальный ток КЗ
                    Try
                        If xlWorkSheet.Range("FM" & jCount.ToString).Value <> Nothing Then .SinglePhShortCurrentMax = MyMod.RoundSign(xlWorkSheet.Range("FM" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                    'Однофазный минимальный ток КЗ
                    Try
                        If xlWorkSheet.Range("FM" & jCount.ToString).Value <> Nothing Then .SinglePhShortCurrentMin = MyMod.RoundSign(xlWorkSheet.Range("FM" & jCount.ToString).Value, 2).ToString
                    Catch ex As Exception
                    End Try

                End With

                If xlWorkSheet.Cells(jCount, 4).mergecells Then

                    Dim mCells As Integer = xlWorkSheet.Cells(jCount, 4).MergeArea.Cells.Count

                    jCount = jCount + mCells - 1

                End If


                fpos = fpos + 1

            Next jCount



            icount = icount + 1
        Next



        xlWorkBook.Close(False)
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)


        'начинаем чертить

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim lock As DocumentLock = acDoc.LockDocument()
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim xCoord As Integer = 0
        Dim yCoord As Integer = 0


        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)



            For icount = 0 To DBcount - 1

                'Вставляем стартовый блок
                If acBlkTbl.Has("Table1") Then

                    'Определяем номер блока
                    Dim blID As ObjectId = acBlkTbl("Table1")
                    'Указваем точку вставки

                    Dim inspoint As Point3d = New Point3d(xCoord, yCoord, 0)
                    'Вставляем блок  

                    Dim br As BlockReference = New BlockReference(inspoint, blID)
                    'запись

                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    'сохранение изменений
                    acBlkTblRec.AppendEntity(br)
                    acTrans.AddNewlyCreatedDBObject(br, True)

                Else

                    acEd.WriteMessage("Не найден блок Table1" & vbCrLf)

                End If

                If acBlkTbl.Has("Start") Then

                    'Определяем номер блока
                    Dim blID As ObjectId = acBlkTbl("Start")

                    'Указваем точку вставки
                    Dim inspoint As Point3d = New Point3d(xCoord + 50, yCoord + 130, 0)

                    'Вставляем блок  
                    Dim br As BlockReference = New BlockReference(inspoint, blID)

                    'запись
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    'сохранение изменений
                    acBlkTblRec.AppendEntity(br)
                    acTrans.AddNewlyCreatedDBObject(br, True)

                    'вставляем атрибуты

                    acBlkTblRec = blID.GetObject(OpenMode.ForRead)
                    For Each objid As ObjectId In acBlkTblRec
                        Dim obj As DBObject = objid.GetObject(OpenMode.ForRead)
                        If TypeOf obj Is AttributeDefinition Then
                            Dim ad As AttributeDefinition = objid.GetObject(OpenMode.ForRead)

                            With MyDB(icount)
                                If ad.Tag = "PHASE" And .Phase <> "" Then
                                    Dim Phase As String
                                    If .Phase.ToLower = "3ph" Then
                                        Phase = "A,B,C"
                                    Else
                                        Phase = .Phase
                                    End If
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = Phase
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If
                            End With
                        End If
                    Next
                Else
                    acEd.WriteMessage("Не найден блок Start" & vbCrLf)
                End If

                If acBlkTbl.Has("DBData") Then
                    'Определяем номер блока
                    Dim blID As ObjectId = acBlkTbl("DBData")
                    'Указваем точку вставки
                    Dim inspoint As Point3d = New Point3d(xCoord + 250, yCoord + 180, 0)
                    'Вставляем блок  
                    Dim br As BlockReference = New BlockReference(inspoint, blID)

                    'запись
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    'сохранение изменений
                    acBlkTblRec.AppendEntity(br)
                    acTrans.AddNewlyCreatedDBObject(br, True)


                    'вставляем атрибуты

                    acBlkTblRec = blID.GetObject(OpenMode.ForRead)
                    For Each objid As ObjectId In acBlkTblRec
                        Dim obj As DBObject = objid.GetObject(OpenMode.ForRead)
                        If TypeOf obj Is AttributeDefinition Then
                            Dim ad As AttributeDefinition = objid.GetObject(OpenMode.ForRead)

                            With MyDB(icount)

                                If ad.Tag = "DBNAME" And .dbName <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .dbName
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "INSTALLEDACTIVEPOWER" And .InstalledActivePowerMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .InstalledActivePowerMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "MAXFACTOR" And .MaxFactor <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .MaxFactor
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDACTIVEPOWER" And .RatedActivePower <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedActivePower
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDACTIVEPOWERMAX" And .RatedActivePowerMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedActivePowerMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDREACTIVEPOWER" And .RatedReactivePower <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedReactivePower
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDREACTIVEPOWERMAX" And .RatedReactivePowerMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedReactivePowerMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDFULLPOWER" And .RatedFullPower <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedFullPower
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDFULLPOWERMAX" And .RatedFullPowerMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedFullPowerMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDCURRENT" And .RatedCurrent <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedCurrent
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RATEDCURRENTMAX" And .RatedCurrentMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedCurrentMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "POWERFACTOR" And .PowerFactor <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .PowerFactor
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "DPY" And .InstalledPowerNonSym <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .InstalledPowerNonSym
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "DIR" And .RatedCurrentNonSym <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RatedCurrentNonSym
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "VOLTAGE" And .FullVoltageDeviation <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .FullVoltageDeviation
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CURRENT3" And .ThreePhShortCurrentMax <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .ThreePhShortCurrentMax
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CURRENT1" And .SinglePhShortCurrentMin <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .SinglePhShortCurrentMin
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If


                            End With


                        End If
                    Next

                Else
                    acEd.WriteMessage("Не найден блок DBData" & vbCrLf)
                End If




                If acBlkTbl.Has("sw6") Then
                    'Определяем номер блока
                    Dim blID As ObjectId = acBlkTbl("sw6")
                    'Указваем точку вставки
                    Dim inspoint As Point3d = New Point3d(xCoord + 75, yCoord + 130, 0)
                    'Вставляем блок  
                    Dim br As BlockReference = New BlockReference(inspoint, blID)

                    'запись
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    'сохранение изменений
                    acBlkTblRec.AppendEntity(br)
                    acTrans.AddNewlyCreatedDBObject(br, True)

                    'вставляем атрибуты
                    acBlkTblRec = blID.GetObject(OpenMode.ForRead)
                    For Each objid As ObjectId In acBlkTblRec
                        Dim obj As DBObject = objid.GetObject(OpenMode.ForRead)
                        If TypeOf obj Is AttributeDefinition Then
                            Dim ad As AttributeDefinition = objid.GetObject(OpenMode.ForRead)

                            With MyDB(icount)

                                If ad.Tag = "PHASE" And .Phase <> "" Then
                                    Dim Phase As String
                                    If .Phase.ToLower = "3ph" Then
                                        Phase = "A,B,C"
                                    Else
                                        Phase = .Phase
                                    End If

                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = Phase.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "NUMBER" And .Number <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .Number.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHTYPE" And .SwitchType <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .SwitchType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHRELEASE" And .SwitchRelease <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).SwitchRelease.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHNOMINALCURRENT" And .SwitchNominalCurrent <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .SwitchNominalCurrent.ToString & "А"
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHRELEASECURRENT" And .SwitchReleaseCurrent <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .SwitchReleaseCurrent.ToString & "А"
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RCDTYPE" And .RCDType <> "" And .RCDType <> "-" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RCDType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RCDSETPOINT" And .RCDSetPoint <> "" And .RCDSetPoint <> "-" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = .RCDSetPoint.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                Dim LINE_A As String = " "

                                Try
                                    If cbLabel.Checked Then LINE_A = LINE_A + .Label
                                Catch Ex As Exception
                                End Try

                                Try
                                    If cbRatedActivePower.Checked Then LINE_A = LINE_A + "-" + .RatedActivePower
                                Catch Ex As Exception
                                End Try

                                Try
                                    If cbPowerFactor.Checked Then LINE_A = LINE_A + "-" + .PowerFactor
                                Catch Ex As Exception
                                End Try

                                Try
                                    If cbRatedCurrent.Checked Then LINE_A = LINE_A + "-" + .RatedCurrent
                                Catch Ex As Exception
                                End Try

                                Try
                                    If cbEstimatedLength.Checked Then LINE_A = LINE_A + "-" + .EstimatedLength
                                Catch Ex As Exception
                                End Try

                                Dim LINE_B As String = " "

                                Try
                                    If cbCableVoltageDeviation.Checked Then LINE_B = LINE_B + .CableVoltageDeviation & "%"
                                Catch Ex As Exception
                                End Try

                                Try
                                    If cbCable.Checked Then LINE_B = LINE_B + "-" & .Cable
                                Catch Ex As Exception
                                End Try

                                If ad.Tag = "LINE_A" And LINE_A <> "" And .Name.ToLower <> "резерв" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = LINE_A
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "LINE_B" And LINE_B <> "" And .Name.ToLower <> "резерв" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = LINE_B
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If
                            End With
                        End If
                    Next


                    With MyDB(icount)

                        Dim vState As String = "SWITCH_1PH"

                        If .Phase <> "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_1PH"

                        'Выключатель трехфазное
                        If .Phase = "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_3PH"

                        'Выключатель однофазный + УЗО
                        If .Phase <> "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_RCD_1PH"

                        'Выключатель трехфазный + УЗО
                        If .Phase = "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_RCD_3PH"

                        'Выключатель однофазный + контактор
                        If .Phase <> "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_CONTACTOR_1PH"

                        'Выключатель трехфазное + контактор
                        If .Phase = "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_CONTACTOR_3PH"

                        'Выключатель однофазный + УЗО + контактор
                        If .Phase <> "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_RCD_CONTACTOR_1PH"

                        'Выключатель трехфазный + УЗО + контактор
                        If .Phase = "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_RCD_CONTACTOR_3PH"

                        'Рубильник
                        If .Number <> Nothing Then
                            If InStr(LCase(.Number), "qs", CompareMethod.Text) <> 0 Then
                                If .Phase <> "3PH" Then vState = "BREAKER_1PH"
                                If .Phase = "3PH" Then vState = "BREAKER_3PH"
                            End If
                        End If

                        Dim dynBrefColl As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection
                        For Each dynbrefProps As DynamicBlockReferenceProperty In dynBrefColl
                            If dynbrefProps.PropertyName = "BreakerType" Then
                                dynbrefProps.Value = vState
                            End If
                        Next

                    End With

                Else
                    acEd.WriteMessage("Не найден блок sw6" & vbCrLf)
                End If


                Dim jcount As Integer
                For jcount = 0 To MyDB(icount).FeadersCount

                    If acBlkTbl.Has("sw1") Then
                        'Определяем номер блока
                        Dim blID As ObjectId = acBlkTbl("sw1")
                        'Указваем точку вставки
                        Dim inspoint As Point3d = New Point3d(xCoord + 50 + jcount * 25, yCoord, 0)
                        'Вставляем блок  
                        Dim br As BlockReference = New BlockReference(inspoint, blID)

                        'запись
                        Dim acBlkTblRec As BlockTableRecord
                        acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        'сохранение изменений
                        acBlkTblRec.AppendEntity(br)
                        acTrans.AddNewlyCreatedDBObject(br, True)

                        'вставляем атрибуты
                        acBlkTblRec = blID.GetObject(OpenMode.ForRead)
                        For Each objid As ObjectId In acBlkTblRec
                            Dim obj As DBObject = objid.GetObject(OpenMode.ForRead)
                            If TypeOf obj Is AttributeDefinition Then
                                Dim ad As AttributeDefinition = objid.GetObject(OpenMode.ForRead)

                                If ad.Tag = "PHASE" And MyDB(icount).Feaders(jcount).Phase <> "" Then
                                    Dim Phase As String
                                    If MyDB(icount).Feaders(jcount).Phase.ToLower = "3ph" Then
                                        Phase = "A,B,C"
                                    Else
                                        Phase = MyDB(icount).Feaders(jcount).Phase
                                    End If

                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = Phase.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "NUMBER" And MyDB(icount).Feaders(jcount).Number <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).Number.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHTYPE" And MyDB(icount).Feaders(jcount).SwitchType <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).SwitchType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHRELEASE" And MyDB(icount).Feaders(jcount).SwitchRelease <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).SwitchRelease.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHNOMINALCURRENT" And MyDB(icount).Feaders(jcount).SwitchNominalCurrent <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).SwitchNominalCurrent.ToString & "А"
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "SWITCHRELEASECURRENT" And MyDB(icount).Feaders(jcount).SwitchReleaseCurrent <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).SwitchReleaseCurrent.ToString & "А"
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RCDTYPE" And MyDB(icount).Feaders(jcount).RCDType <> "" And MyDB(icount).Feaders(jcount).RCDType <> "-" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).RCDType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "RCDSETPOINT" And MyDB(icount).Feaders(jcount).RCDSetPoint <> "" And MyDB(icount).Feaders(jcount).RCDSetPoint <> "-" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).RCDSetPoint.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CONTACTOR" And MyDB(icount).Feaders(jcount).Contactor <> "-" And MyDB(icount).Feaders(jcount).Contactor <> "" And MyDB(icount).Feaders(jcount).Contactor <> " " Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).Contactor.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CONTACTOR_RCD" And MyDB(icount).Feaders(jcount).Contactor <> "-" And MyDB(icount).Feaders(jcount).Contactor <> "" And MyDB(icount).Feaders(jcount).Contactor <> " " Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).Contactor.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CONTACTOR_TYPE" And MyDB(icount).Feaders(jcount).ContactorType <> "-" And MyDB(icount).Feaders(jcount).ContactorType <> "" And MyDB(icount).Feaders(jcount).ContactorType <> " " Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).ContactorType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "CONTACTOR_RCD_TYPE" And MyDB(icount).Feaders(jcount).ContactorType <> "-" And MyDB(icount).Feaders(jcount).ContactorType <> "" And MyDB(icount).Feaders(jcount).ContactorType <> " " Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).ContactorType.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "LABEL" And MyDB(icount).Feaders(jcount).PlanLabel <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).PlanLabel.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "INSTALLEDPOWER" And MyDB(icount).Feaders(jcount).InstalledActivePower <> "" Then
                                    If MyDB(icount).Feaders(jcount).Name.ToLower <> "резерв" Then

                                        Dim ar As AttributeReference = New AttributeReference()
                                        ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                        ar.TextString = MyDB(icount).Feaders(jcount).InstalledActivePower.ToString
                                        br.AttributeCollection.AppendAttribute(ar)
                                        acTrans.AddNewlyCreatedDBObject(ar, True)
                                    End If
                                End If

                                If ad.Tag = "CURRENT" And MyDB(icount).Feaders(jcount).RatedCurrent <> "" Then
                                    If MyDB(icount).Feaders(jcount).Name.ToLower <> "резерв" Then
                                        Dim ar As AttributeReference = New AttributeReference()
                                        ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                        ar.TextString = MyDB(icount).Feaders(jcount).RatedCurrent.ToString
                                        br.AttributeCollection.AppendAttribute(ar)
                                        acTrans.AddNewlyCreatedDBObject(ar, True)

                                    End If
                                End If

                                If ad.Tag = "VOLTAGE" And MyDB(icount).Feaders(jcount).Voltage <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).Voltage.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                If ad.Tag = "TEXT" And MyDB(icount).Feaders(jcount).Name <> "" Then
                                    Dim ar As AttributeReference = New AttributeReference()
                                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                    ar.TextString = MyDB(icount).Feaders(jcount).Name.ToString
                                    br.AttributeCollection.AppendAttribute(ar)
                                    acTrans.AddNewlyCreatedDBObject(ar, True)
                                End If

                                With MyDB(icount).Feaders(jcount)



                                    Dim LINE_A As String = " "

                                    Try
                                        If cbLabel.Checked Then LINE_A = LINE_A + .Label
                                    Catch Ex As Exception
                                    End Try

                                    Try
                                        If cbRatedActivePower.Checked Then LINE_A = LINE_A + "-" + .RatedActivePower
                                    Catch Ex As Exception
                                    End Try

                                    Try
                                        If cbPowerFactor.Checked Then LINE_A = LINE_A + "-" + .PowerFactor
                                    Catch Ex As Exception
                                    End Try

                                    Try
                                        If cbRatedCurrent.Checked Then LINE_A = LINE_A + "-" + .RatedCurrent
                                    Catch Ex As Exception
                                    End Try

                                    Try
                                        If cbEstimatedLength.Checked Then LINE_A = LINE_A + "-" + .EstimatedLength
                                    Catch Ex As Exception
                                    End Try

                                    Dim LINE_B As String = " "

                                    Try
                                        If cbCableVoltageDeviation.Checked Then LINE_B = LINE_B + .CableVoltageDeviation & "%"
                                    Catch Ex As Exception
                                    End Try

                                    Try
                                        If cbCable.Checked Then LINE_B = LINE_B + "-" & .Cable
                                    Catch Ex As Exception
                                    End Try




                                    If ad.Tag = "LINE_A" And LINE_A <> "" And .Name <> "" Then
                                        If .Name.ToLower <> "резерв" Then
                                            Dim ar As AttributeReference = New AttributeReference()
                                            ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                            ar.TextString = LINE_A
                                            br.AttributeCollection.AppendAttribute(ar)
                                            acTrans.AddNewlyCreatedDBObject(ar, True)
                                        End If
                                    End If

                                    If ad.Tag = "LINE_B" And LINE_B <> "" And .Name <> "" Then
                                        If .Name.ToLower <> "резерв" Then
                                            Dim ar As AttributeReference = New AttributeReference()
                                            ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                            ar.TextString = LINE_B
                                            br.AttributeCollection.AppendAttribute(ar)
                                            acTrans.AddNewlyCreatedDBObject(ar, True)
                                        End If
                                    End If
                                End With
                            End If
                        Next

                        With MyDB(icount).Feaders(jcount)
                            Dim vState As String = "SWITCH_1PH"

                            If .Phase <> "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_1PH"

                            'Выключатель трехфазное
                            If .Phase = "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_3PH"

                            'Выключатель однофазный + УЗО
                            If .Phase <> "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_RCD_1PH"

                            'Выключатель трехфазный + УЗО
                            If .Phase = "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType = "-" Or .ContactorType = "" Or .ContactorType = " ") Then vState = "SWITCH_RCD_3PH"

                            'Выключатель однофазный + контактор
                            If .Phase <> "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_CONTACTOR_1PH"

                            'Выключатель трехфазный + контактор
                            If .Phase = "3PH" And (.RCDSetPoint = "-" Or .RCDSetPoint = "" Or .RCDSetPoint = " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_CONTACTOR_3PH"

                            'Выключатель однофазный + УЗО + контактор
                            If .Phase <> "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_RCD_CONTACTOR_1PH"

                            'Выключатель трехфазный + УЗО + контактор
                            If .Phase = "3PH" And (.RCDSetPoint <> "-" And .RCDSetPoint <> "" And .RCDSetPoint <> " ") And (.ContactorType <> "-" And .ContactorType <> "" And .ContactorType <> " ") Then vState = "SWITCH_RCD_CONTACTOR_3PH"






                            Dim dynBrefColl As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection
                            For Each dynbrefProps As DynamicBlockReferenceProperty In dynBrefColl
                                If dynbrefProps.PropertyName = "BreakerType" Then
                                    dynbrefProps.Value = vState
                                End If
                            Next
                        End With

                    Else
                        acEd.WriteMessage("Не найден блок sw1" & vbCrLf)
                    End If

                    If acBlkTbl.Has("CONSUMER") Then
                        'Определяем номер блока
                        Dim blID As ObjectId = acBlkTbl("CONSUMER")
                        'Указваем точку вставки
                        Dim inspoint As Point3d = New Point3d(xCoord + 50 + 25 / 2 + jcount * 25, yCoord + 62.5, 0)
                        'Вставляем блок  
                        Dim br As BlockReference = New BlockReference(inspoint, blID)

                        'запись
                        Dim acBlkTblRec As BlockTableRecord
                        acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        'сохранение изменений
                        acBlkTblRec.AppendEntity(br)
                        acTrans.AddNewlyCreatedDBObject(br, True)

                        With MyDB(icount).Feaders(jcount)
                            Dim vState As String = "CableOutlet"
                            If .Name <> "" Then
                                If InStr(.Name.ToString.ToLower, "освещ") <> 0 Then vState = "FluorescentLamp"
                                If InStr(.Name.ToString.ToLower, "освет") <> 0 Then vState = "FluorescentLamp"
                                If InStr(.Name.ToString.ToLower, "розет") <> 0 Then vState = "Socket"
                                If InStr(.Name.ToString.ToLower, "эвакуац") <> 0 Then vState = "IncandescentLamp"
                                If InStr(.Name.ToString.ToLower, "щит") <> 0 Then vState = "DBoard"
                            End If

                            Dim dynBrefColl As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection
                            For Each dynbrefProps As DynamicBlockReferenceProperty In dynBrefColl
                                If dynbrefProps.PropertyName = "ConsumerType" Then
                                    dynbrefProps.Value = vState
                                End If
                            Next
                        End With

                    Else
                        acEd.WriteMessage("Не найден блок CONSUMER" & vbCrLf)
                    End If










                Next jcount
                yCoord = yCoord - 250
            Next icount

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()

        End Using

        lock.Dispose()
        'Закрываем
        Me.Close()
        MsgBox("Схема сгенерирована!!!")
        acEd.WriteMessage("Схема сгенерирована!!!" & vbCrLf)

    End Sub

    Private Sub DataFileName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataFileName.TextChanged

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
