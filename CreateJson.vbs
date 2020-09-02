Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

set json = CreateObject("Chilkat_9_5_0.JsonObject")

'  The only reason for failure in the following lines of code would be an out-of-memory condition..

'  An index value of -1 is used to append at the end.
index = -1

success = json.AddStringAt(-1,"Title","Pan's Labyrinth")
success = json.AddStringAt(-1,"Director","Guillermo del Toro")
success = json.AddStringAt(-1,"Original_Title","El laberinto del fauno")
success = json.AddIntAt(-1,"Year_Released",2006)

json.EmitCompact = 0
outFile.WriteLine(json.Emit())

outFile.Close