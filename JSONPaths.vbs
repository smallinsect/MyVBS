Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

set json = CreateObject("Chilkat_9_5_0.JsonObject")
json.EmitCompact = 0

' Assume the file contains the data as shown above..
success = json.LoadFile("qa_data/json/pathSample.json")
If (success <> 1) Then
    outFile.WriteLine(json.LastErrorText)
    WScript.Quit
End If

' First, let's get the value of "cc1"
' The path to this value is: nestedObject.aaa.bb1.cc1
outFile.WriteLine(json.StringOf("nestedObject.aaa.bb1.cc1"))

' Now let's get number 18 from the nestedArray.
' It is located at nestedArray[1][2][1]
' (remember: Indexing is 0-based)
outFile.WriteLine("This should be 18: " & json.IntOf("nestedArray[1][2][1]"))

' We can do the same thing in a more roundabout way using the 
' I, J, and K properties.  (The I,J,K properties will be convenient
' for iterating over arrays, as we'll see later.)
json.I = 1
json.J = 2
json.K = 1
outFile.WriteLine("This should be 18: " & json.IntOf("nestedArray[i][j][k]"))

' Let's iterate over the array containing the numbers 17, 18, 19, 20.
' First, use the SizeOfArray method to get the array size:
sz = json.SizeOfArray("nestedArray[1][2]")
' The size should be 4.
outFile.WriteLine("size of array = " & sz & " (should equal 4)")

' Now iterate...

For i = 0 To sz - 1
    json.I = i
    outFile.WriteLine(json.IntOf("nestedArray[1][2][i]"))
Next

' Let's use a triple-nested loop to iterate over the nestedArray:

' szI should equal 1.
szI = json.SizeOfArray("nestedArray")
For i = 0 To szI - 1
    json.I = i

    szJ = json.SizeOfArray("nestedArray[i]")
    For j = 0 To szJ - 1
        json.J = j

        szK = json.SizeOfArray("nestedArray[i][j]")
        For k = 0 To szK - 1
            json.K = k

            outFile.WriteLine(json.IntOf("nestedArray[i][j][k]"))
        Next
    Next
Next

' Now let's examine how to navigate to JSON objects contained within JSON arrays.
' This line of code gets the value "kiwi" contained within "mixture"
outFile.WriteLine(json.StringOf("mixture.arrayA[2].fruit"))

' This line of code gets the color "yellow"
outFile.WriteLine(json.StringOf("mixture.arrayA[1].colors[0]"))

' Getting an object at a path:
' This gets the 2nd object in "arrayA"
' obj2 is a Chilkat_9_5_0.JsonObject
Set obj2 = json.ObjectOf("mixture.arrayA[1]")
' This object's "animal" should be "plankton"
outFile.WriteLine(obj2.StringOf("animal"))

' Note that paths are relative to the object, not the absolute root of the JSON document.
' Starting from obj2, "purple" is at "colors[2]"
outFile.WriteLine(obj2.StringOf("colors[2]"))

' Getting an array at a path:
' This gets the array containing the colors red, green, blue:
' arr1 is a Chilkat_9_5_0.JsonArray
Set arr1 = json.ArrayOf("mixture.arrayA[0].colors")
szArr1 = arr1.Size
For i = 0 To szArr1 - 1
    outFile.WriteLine(i & ": " & arr1.StringAt(i))
Next

' The Chilkat JSON path uses ".", "[", and "]" chars for separators.  When a name
' contains one of these chars, use double-quotes in the path:
outFile.WriteLine(json.StringOf("""name.with.dots"".grain"))

outFile.Close