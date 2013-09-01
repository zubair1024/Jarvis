Imports System.Speech

Public Class Form1

    Dim WithEvents reco As New Recognition.SpeechRecognitionEngine

    Private Sub SetColor(ByVal color As System.Drawing.Color)

        Dim synth As New Synthesis.SpeechSynthesizer

        synth.SpeakAsync("setting the back color to " + color.ToString)

        Me.BackColor = color

    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        reco.SetInputToDefaultAudioDevice()

        Dim gram As New Recognition.SrgsGrammar.SrgsDocument

        Dim colorRule As New Recognition.SrgsGrammar.SrgsRule("color")

        Dim colorsList As New Recognition.SrgsGrammar.SrgsOneOf("red", "green", "blue")

        colorRule.Add(colorsList)

        gram.Rules.Add(colorRule)

        gram.Root = colorRule

        reco.LoadGrammar(New Recognition.Grammar(gram))

        reco.RecognizeAsync()

    End Sub

    Private Sub reco_RecognizeCompleted(ByVal sender As Object, ByVal e As System.Speech.Recognition.RecognizeCompletedEventArgs) Handles reco.RecognizeCompleted

        reco.RecognizeAsync()

    End Sub

    Private Sub reco_SpeechRecognized(ByVal sender As Object, ByVal e As System.Speech.Recognition.RecognitionEventArgs) Handles reco.SpeechRecognized

        Select Case e.Result.Text

            Case "red"

                SetColor(Color.Red)

            Case "green"

                SetColor(Color.Green)

            Case "blue"

                SetColor(Color.Blue)

        End Select

    End Sub

End Class