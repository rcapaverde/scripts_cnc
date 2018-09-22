option explicit

Dim fso
Dim inputFile
Dim outputFile

Dim nomeEntrada
Dim feedRate
Dim freeHeight
Dim drillDepth

feedRate = 80
freeHeight = 4
drillDepth = -3

if WScript.Arguments.count = 0 then
    WScript.Echo "Uso: <arquivo.dri>"
    WScript.Quit
end if

nomeEntrada = WScript.Arguments(0)

Dim iArg
for iArg = 1 to WScript.Arguments.count - 1
    Dim arg
    arg = WScript.Arguments(iArg)
next

Set fso = CreateObject("Scripting.FileSystemObject")
Set inputFile = fso.OpenTextFile(nomeEntrada, 1, false, 0)
Set outputFile = fso.OpenTextFile(nomeEntrada + ".txt", 2, true, 0)

Dim line
Dim toolDiameters(100)

do while not inputFile.AtEndOfStream
    line = inputFile.ReadLine

    if mid(line, 1, 1) = "X" then
        Dim x
        Dim y
        Dim separador

        separador = instr(line, "Y")
        x = cdbl(mid(line, 2, separador - 2)) * 25.4 / 100000 * -1
        y = cdbl(mid(line, separador + 1)) * 25.4 / 100000 * -1

        outputFile.WriteLine pontoDecimal(cstr(x)) & chr(9) & pontoDecimal(cstr(y))

    elseif mid(line, 1, 1) = "%" then
        outputFile.WriteLine line

    elseif mid(line, 1, 1) = "T" then
        Dim tool
        Dim diameter
        separador = instr(line, "C")
        if separador > 0 then
            tool = clng(mid(line, 2, separador - 2))
            diameter = cdbl(virgulaDecimal(mid(line, separador + 1))) * 25.4
            toolDiameters(tool) = diameter
            outputFile.WriteLine "% " & line
        else
            tool = clng(mid(line, 2))
            outputFile.WriteLine ""
            outputFile.WriteLine "% *****************************"
            outputFile.WriteLine "% " & line & " (" & FormatNumber(toolDiameters(tool), 2) & " mm)"
        end if

    else
        outputFile.WriteLine "% " & line
    end if
Loop

inputFile.Close
outputFile.Close

function pontoDecimal(str)
    dim v
    v = instr(str, ",")
    if v > 0 then
        str = mid(str, 1, v - 1) & "." & mid(str, v + 1)
    end if
    pontoDecimal = str
end function

function virgulaDecimal(str)
    dim v
    v = instr(str, ".")
    if v > 0 then
        str = mid(str, 1, v - 1) & "," & mid(str, v + 1)
    end if
    virgulaDecimal = str
end function
