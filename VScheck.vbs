Option Explicit
Dim cnt, i, main_path, file_name, stPath, WshShell, FSO, FD, obj

' �E�B���X�Z�L�����e�B�̎��s
main_path = "C:\Program Files (x86)\K7 Computing\K7TSecurity\"
file_name = array("K7TSUpdT.exe","K7AVScan.exe","K7TSMain.exe") 

cnt = ubound(file_name)	' �z��̐��i���[�v�񐔁j
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

for i = 0 to cnt

	Set FD = FSO.GetFolder(main_path)
	For Each obj in FD.files
	If obj.name = file_name(i) then

		'WScript.Echo stPath & vbcrlf & obj.name�@'�t�@�C���̑��݊m�F���o������p�X�\��
		Call WshShell.Run( """" & main_path & file_name(i) & """", 0, True )	 '�N��
		exit for

	end if
	set FD = Nothing
	next	'obj

next 'i

Set FSO = Nothing
Set WshShell = Nothing
