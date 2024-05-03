Option Explicit

' ありがとう ひどすぎるコード！！ 

Dim intArgsCount

intArgsCount = Wscript.Arguments.Count

' 引数チェック
if (intArgsCount = 0) then

    msgbox "mp3 Drag and Drop"
end if


Dim ret
Dim objShell, objFS, objFolder, objFolderItems , objFolderItem, objFile
Dim cnt
Dim filename
Dim foldername
Dim tag , tagFilename
Dim tarfilepath

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("Shell.Application")

for cnt = 0 to intArgsCount - 1 step 1

    ' msgbox Wscript.Arguments(cnt)
    tarfilepath = Wscript.Arguments(cnt)
    ret = objFS.FileExists(tarfilepath)

    '
    if ((ret = true) AND _
        (objFS.GetExtensionName(tarfilepath) = "mp3")) then


        filename = objFS.getFileName(tarfilepath)
        foldername = objFS.GetParentFolderName(tarfilepath)



        Set objFolder = objShell.Namespace(foldername)
        Set objFolderItems = objFolder.Items
        Set objFolderItem = objFolderItems.Item(filename)

        tag = objFolder.GetDetailsOf(objFolderItem, 21)

        tagFilename = foldername + "\" + tag + ".mp3"
        ret = objFS.FileExists(tagFilename)

        ' changed filename already exists
        if (ret = true) then

            msgbox "File:" + tagFilename + " is exist"
        else

            Set objFile = objFS.GetFile(tarfilepath)
            objFile.Name = tag + ".mp3"
            Set objFile = Nothing
        end if

        Set objFolder = Nothing
        Set objFolderItems = Nothing
        Set objFolderItem = Nothing
    end if
next

Set objFS = Nothing
Set objShell = Nothing
