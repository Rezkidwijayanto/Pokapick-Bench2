Module Module1

    Public Class ResizeFormClass
        'Original form width.
        Private Shared m_FormWidth As Long
        Private Shared m_FormHeight As Long



        Public Shared Sub SubResize(ByVal F As Form, ByVal percentW As Double, ByVal percentH As Double)
            Dim FormHeight As Long
            Dim FormWidth As Long
            Dim HeightChange As Double, WidthChange As Double



            Call SaveInitialStates(F)




            'Calculate the new height and width the form needs to be resized to, based on the current avaible screen area.
            FormHeight = Int((Screen.PrimaryScreen.WorkingArea.Height) * (percentH / 100))
            FormWidth = Int((Screen.PrimaryScreen.WorkingArea.Width) * (percentW / 100))



            'Use the Form that is to be resized.
            With F
                'Change the demensions and position of the form.

                .Height = FormHeight
                .Width = FormWidth

                HeightChange = .ClientSize.Height / m_FormHeight
                WidthChange = .ClientSize.Width / m_FormWidth

            End With
            'Calculate ratio current avaible screen area/form size

            'Notify the class that the form has been resized.
            SubChangeWithRatio(F, WidthChange, HeightChange)

        End Sub

        Private Shared Sub SaveInitialStates(ByVal F As Form)

            'Use the form that is being resized.
            With F
                'Check if the form is a MDI form.

                'Set the FormWidth and FormHeight variables to the Form's Scale Width and Height.
                m_FormWidth = .ClientSize.Width
                m_FormHeight = .ClientSize.Height

            End With

        End Sub


        Public Shared Sub SubChangeWithRatio(ByVal F As Form, ByVal RapportoW As Single, ByVal RapportoH As Single)
            'uses a recursive routine
            For Each ctl As Control In F.Controls
                ResizeControlAndIncludedControls(ctl, RapportoW, RapportoH)
            Next

        End Sub

        Private Shared Sub ResizeControlAndIncludedControls(ByRef ctl As Control, ByVal RapportoW As Single, ByVal RapportoH As Single)



            Dim ChildCtl As Control

            For Each ChildCtl In ctl.Controls

                ResizeControlAndIncludedControls(ChildCtl, RapportoW, RapportoH)

            Next
            ResizeControl(ctl, RapportoW, RapportoH)
        End Sub

        Private Shared Sub ResizeControl(ByRef ctl As Control, ByVal RapportoW As Single, ByVal RapportoH As Single)
            Dim lb As New ListBox, intlH As Boolean
            Try
                If TypeOf ctl Is ListBox Then


                    lb = CType(ctl, ListBox)
                    intlH = lb.IntegralHeight
                    lb.IntegralHeight = False

                    ctl.Left = CInt(ctl.Left * RapportoW)
                    ctl.Top = CInt(ctl.Top * RapportoH)
                    ctl.Width = CInt(ctl.Width * RapportoW)
                    ctl.Height = CInt(ctl.Height * RapportoH)

                Else

                    ctl.Left = CInt(ctl.Left * RapportoW)
                    ctl.Top = CInt(ctl.Top * RapportoH)
                    ctl.Width = CInt(ctl.Width * RapportoW)
                    ctl.Height = CInt(ctl.Height * RapportoH)

                End If

                lb.IntegralHeight = intlH
                If TypeOf ctl Is ListView Then
                    Try
                        ResizeColumns(ctl, RapportoW, RapportoH)
                    Catch ex As Exception
                    End Try
                End If
                Try
                    ResizeControlFont(ctl, RapportoW, RapportoH)

                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try

        End Sub

        Private Shared Sub ResizeControlFont(ByRef Ct As Control, ByVal RapportoW As Single, ByVal RapportoH As Single)

            'Resizes the control font and, in the case of some controls, as the listview
            ' resizes the columns also

            Try

                Dim FSize As Double = Ct.Font.Size
                Dim FStile As FontStyle = Ct.Font.Style
                Dim FNome As String = Ct.Font.Name
                Dim NuovoSize As Double = FSize



                NuovoSize = FSize * Math.Sqrt(RapportoW * RapportoH)
                Dim NFont As New Font(FNome, CSng(NuovoSize), FStile)
                Ct.Font = NFont

            Catch

            End Try

        End Sub

        Private Shared Sub ResizeColumns(ByRef ct As Control, ByVal RapportoW As Single, ByVal RapportoH As Single)

            Dim c As ColumnHeader
            For Each c In CType(ct, ListView).Columns
                c.Width = CInt(c.Width * RapportoW)
            Next

        End Sub
    End Class
End Module