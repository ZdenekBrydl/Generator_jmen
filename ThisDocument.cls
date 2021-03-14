VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
Private Slabiky()

Sub AutoOpen()
  Dim I%

  Slabiky = Array(Array(Array("Ar", "Nor", "Gor", "Tab", "Zen", "Ba", "Run", "Ken", "Mat", "Pro", "Mil", "Lan", "Len", "Far", "Zar"), _
                  Array("ros", "nig", "nul", "nag", "ner", "nor", "tor", "ul", "rin", "us", "near", "tel", "fos", "nur", "tan"), _
                  Array("an", "as", "in", "en", "is", "es", "un", "at", "ar", "os", "or", "er", "on", "ir", "ad")), _
            Array(Array("Vet", "Ker", "Ag", "Mi", "Na", "Ka", "Fer", "Ri", "Eg", "Pen"), _
                  Array("a", "oa", "ta", "ea", "ia", "ria", "lia", "la", "loa", "lea"), _
                  Array("ria", "nia", "lia", "sia", "mia", "kia", "ia", "hia", "cha", "ta")), _
            Array(Array("Jan", "Ar", "Bos", "Bar", "Gis", "Tos", "Far", "Ul", "Xe", "Bor", "Nar", "Nor", "Ner", "Moj", "Mej"), _
                  Array("ir", "gir", "sir", "is", "in", "il", "or", "an", "im", "am", "tis", "ris", "ois", "eis", "eir"), _
                  Array("in", "is", "ir", "il", "im", "ig", "an", "am", "ar", "un", "el", "er", "on", "om", "en")))

  Randomize
  Parametry.Show
  
  For I = 1 To Parametry.Jmen
    Selection.TypeText NahodnaSlabika(Parametry.Rod.ListIndex, 0)
    If Parametry.Slabik.ListIndex > 0 Then Selection.TypeText NahodnaSlabika(Parametry.Rod.ListIndex, 1)
    If Parametry.Slabik.ListIndex > 1 Then Selection.TypeText NahodnaSlabika(Parametry.Rod.ListIndex, 2)
    Selection.TypeText vbNewLine
  Next

End Sub

Function NahodnaSlabika(Rod%, Sl%)
  NahodnaSlabika = Slabiky(Rod)(Sl)(Int((UBound(Slabiky(Rod)(Sl)) + 1) * Rnd))
End Function
