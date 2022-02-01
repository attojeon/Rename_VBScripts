' 숫자로 시작하는 파일의 번호를 1씩 증가시킨다.
' source 형식 : 23.문자~.py 
' dest 형식 : 24.문자~.py



Option Explicit
Function OnGetNewName(ByRef getNewNameData)
	' 여기에 스크립트 코드 추가.
	' 
	' 입력 (모두 읽기 전용):
	' 
	' - getNewNameData.item:
	'     항목에 대한 정보가 있는 오브젝트가 이름 변경됨.
	'     e.g. item.name, item.name_stem, item.ext,
	'          item.is_dir, item.size, item.modify,
	'          item.metadata, etc.
	'     item.path는 부모 폴더로의 경로.
	'     item.realpath는 이름을 포함한 파일으로의 전체 경로,
	'          and with things like Collections and Libraries resolved to
	'          the real directories that they point to.
	' - getNewNameData.oldname:
	'     이름 변경 창의 "이전 이름"칸의 내용. 스크립트는
	'     보통 이 항목을 사용 하지 않습니다. 파일의 현재 이름은 item.name을 사용 하세요.
	' - getNewNameData.newname
	'      비 스크립트 측면 으로 제안된 '이름 변경' 창
	'     항목의 새 이름. 스크립트는 보통 item.name 보다는
	'     이 설정으로 잘 동작 합니다..
	' 
	' Return values:
	' 
	' - OnGetNewName=True:
	'     이름 변경 불가능.
	'     The proposed getNewNameData.newname is not used.
	' - OnGetNewName=False: (Default)
	'     이름 변경 가능.
	'     The proposed getNewNameData.newname is used as-is.
	' - OnGetNewName="string"
	'     이름 변경 가능.
	'     파일의 새로운 이름은 스크립트가 반환 하는 스트링 값 입니다.

	dim item, olds, new_name, new_num, i
	Set item = getNewNameData.item
	olds = Split(item.name_stem, ".")
	olds(0) = CStr( CInt(olds(0)+1) ) 
	For i=1 To UBound(olds)
		olds(i) = olds(i)
	Next
	new_name = Join(olds, ".")
	OnGetNewName = new_name & item.ext
End Function
