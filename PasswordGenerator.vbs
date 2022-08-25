Set objFSO=CreateObject("Scripting.FileSystemObject")

if objFSO.FileExists("./new-passcode.txt") Then
objFSO.DeleteFile "./new-passcode.txt"
MsgBox"Welcome To My Password Generator: generating... ",0,"Maxxx Pass Gen"
For x=200 To 300
    Randomize
    vChar = Int(89*Rnd) + 33
    Rndz = Rndz & Chr(vChar)
  Next
outFile = "./new-passcode.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write Rndz & vbCrLf
objFile.Close
else
MsgBox"Welcome To My Password Generator: generating... ",0,"Maxxx Pass Gen"
For x=200 To 300
    Randomize
    vChar = Int(89*Rnd) + 33
    Rndz = Rndz & Chr(vChar)
  Next
outFile = "./new-passcode.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write Rndz & vbCrLf
objFile.Close
End if