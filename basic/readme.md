
# [기초.1. 매크로처음사용.xlsm](https://github.com/KimDaeHo26/vba/raw/main/basic/%EA%B8%B0%EC%B4%88.1.%20%EB%A7%A4%ED%81%AC%EB%A1%9C%EC%B2%98%EC%9D%8C%EC%82%AC%EC%9A%A9.xlsm)
  ## [개발도구 메뉴 표시]
  * 1. [Office 단추] - [Excel 옵션]을 클릭
  * 2. [Excel 옵션] 대화 상자가 나타나면 [기본 설정] 항목의 [리본 메뉴에 개발 도구 탭 표시]를 선택하여 체크 표시한 후 [확인]을 클릭
 
  ## [매크로기록]
  * 1. [개발도구 메뉴] - [매크로 기록] 클릭
  * 2. 엑셀 에서 기록할 작업 실행
  * 3. [기록중지] 클릭
  * 4. [Visual Basic] 클릭 하여 매크로 소스 보기 / 수정
  
  ## [매크로연결]
  * 1. [삽입 메뉴] - [도형] 클릭 하여 도형을 삽입
  * 2. 삽입된 도형 [마우스 우클릭] - [매크로 지정] 클릭 하여 만들어진 매크로 선택 후 [확인] 버튼 클릭
 
  ## [프로젝트암호설정]
  * 1. [프로젝트] 우클릭 - [VBAProject 속성] 클릭
  * 2. [보호] 텝 [읽기 전용으로 프로젝트 잠금 check 암호 입력후 [확인] 버튼 클릭
 
  ## [디버깅]
  * 1. 멈추고 싶은 라인 선택 - [디버그] 메뉴 - [중단점 설정/해제]
  * 2. [Sub/사용자 정의 폼 실행] 클릭
      중단점 설정한 곳에서 멈춤
  * 3. 값 확인 하고 싶은 [변수] 마우스 우클릭 - [조사식 추가]
  * 4. 아래 조사식 부분에서 값 확인 가능
 
# [기초.2. WORKBOOK이벤트.xlsm](https://github.com/KimDaeHo26/vba/raw/main/basic/%EA%B8%B0%EC%B4%88.2.%20WORKBOOK%EC%9D%B4%EB%B2%A4%ED%8A%B8.xlsm)
  ## [워크북이벤트]
  * 1. [ThisWorkbook]더블클릭-[개체] Workbook 선택, [프로시저] Open 선택
  * 2. 프로시저 내에 코딩
        MsgBox ThisWorkbook.Name & "파일이 열렸습니다."
  * 3. 파일을 다시 열었을때 메시지가 표시됨
  * 4. 이벤트 설명은 도움말 참고
 
  ## [워크시트이벤트]
  * 1. [Sheet(N)]더블클릭-[개체] Worksheet 선택, [프로시저] Activate 선택
  * 2. 프로시저 내에 코딩
     MsgBox ActiveSheet.Name
  * 3. [Sheet(N)] 시트가 활성화 되면 메시지가 표시됨
  * 4. 이벤트 설명은 도움말 참고

# [기초.3. 프로시저및에러처리.xlsm](https://github.com/KimDaeHo26/vba/raw/main/basic/%EA%B8%B0%EC%B4%88.3.%20%ED%94%84%EB%A1%9C%EC%8B%9C%EC%A0%80%EB%B0%8F%EC%97%90%EB%9F%AC%EC%B2%98%EB%A6%AC.xlsm)
  ## [프로시저]
  * ByVal, ByRef 개념
  * sub에서 function 호출
  * cell에 사용자 생성 function 사용
 
  ## [에러처리]
  * 에러나는 경우 처리 방법
  * 에러가 나도 다음 계속 진행 : On Error Resume Next
  * 지정된 행으로 무조건 분기 : GoTo line

# [기초.4. 복사.xlsm](https://github.com/KimDaeHo26/vba/raw/main/basic/%EA%B8%B0%EC%B4%88.4.%20%EB%B3%B5%EC%82%AC.xlsm)
  ## [복사]
  * 셀복사 capy, paste
  * 셀복사 value = value
  * 셀복사  FormulaLocal = FormulaLocal
  * 셀복사  새로운 엑셀파일에
  * 셀복사  새로운 엑셀파일에 배열로
 
  ## [이름관리자사용]
  * 셀주소 대신 이름 사용

# [기초.5. 시트암호걸기.xlsm](https://github.com/KimDaeHo26/vba/raw/main/basic/%EA%B8%B0%EC%B4%88.5.%20%EC%8B%9C%ED%8A%B8%EC%95%94%ED%98%B8%EA%B1%B8%EA%B8%B0.xlsm)
  ## [시트암호]
  * 시트암호걸어 특정 셀만 수정가능하게 처리
 
  ## [프로젝트암호]
  * vba 소스 보지 못하게 처리
