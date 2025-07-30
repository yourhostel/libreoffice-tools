REM  *****  BASIC  *****

' MainVariable.bas

Public Const HOSTEL                     = "Саперка"
Public Const MANAGER                    = "Черкашина Світлана Віталіївна"
Public Const TOP_MANAGER                = "top"
Public Const ADMIN_MANAGER              = "admin"
Public Const ENCASH                     = "інкасація"
Public Const BALANCE                    = "баланс"
Public Const PATH_TO_LOGO               = "file:///home/tysser/Documents/LibreOfficeAssets/logo_1.png"
Public Const BACKGROUND                 = "file:///home/tysser/Documents/LibreOfficeAssets/20.png"
Public Const VALID_DURATIONS            = "1;2;3;4;5;6;7;14;21;28"
Public Const VISIBLE_ROWS               = 25
Public Const NEGET_RULES                = "F09D8C98CE"
Public Const LIST_OF_CODES              = "1;2;3;4;5;6;8;9;10;11"
Public Const LIST_OF_CANCELLATIONS      = "Скасування;Пауза;Перенесення;Часткова оплата"
Public Const EXCLUDED_CODES             = " 28 20 21 22 23 7 30 " ' '28' ч/с, '20 +' скасовано, '7' инкасація, '30' баланс
Public Const MAX_SEARCH_RANGE_IN_PRICE  = 28
Public Const MAX_SEARCH_RANGE_IN_ADMINS = 30
Public Const ACTION_CREATE              = "create"
Public Const ACTION_EDIT                = "edit"
Public Const ACTION_CHECK_ROW           = "check_row"
Public Const ACTION_CHECK_CANCEL        = "check_cancel"
Public Const SEARCH_LIST                = "за місцем;по прізвищу;за iм'ям по-батькові;по id;за номером телефону;по адміністратору;за датою заселення;за датою виселення;за датою народження;за терміном;за кодом"
Public Const SEARCH_MONTHS_AGO          = -2 ' від 0 до -12 
Public Const ALL_PLACES                 = "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28"
Public Const ALL_EXPENSES               = "зарплата;повернення;хімія;розхідники;зв’язок;запчастини;обладнання;ремонт;реклама;комуналка;оренда;охорона;податки;ліцензії;штрафи;страховка"
Public Const ALL_INCOMES                = "товари;оренда;послуги;депозити;харчування"
Public Const LIST_OF_HISTORY_FIELDS     = "заселення|прізвище|ім'я по батькові|термін|виселення|сплачено|видаток|прихід|чорний список|створено|причина зсуву|місце|код|id|адмін що редагував"
Public Const TEXT_COLOR                 = "255 255 255"

Function pRGB(s As String) As Long
    Dim arr() As String
    arr = Split(s, " ")
    pRGB = RGB(CInt(arr(0)), CInt(arr(1)), CInt(arr(2)))
End Function

'| Колір           | Значення ("R G B")   |
'| --------------- | -------------------- |
'| 🔴 Червоний     |   "255 0 0"          |
'| 🟢 Зелений      |   "0 255 0"          |
'| 🔵 Синій        |   "0 0 255"          |
'| ⚫ Чорний       |   "0 0 0"            |
'| ⚪ Білий        |   "255 255 255"      |
'| 🟡 Жовтий       |   "255 255 0"        |
'| 🟣 Фіолетовий   |   "255 0 255"        |
'| 🟠 Помаранчевий |   "255 127 0"        |
'| 🩶 Сірий        |   "128 128 128"      |

' ==== Повертає мапу відповідності полів форми до назв колонок таблиці ====
Function GetFieldToColumnMap() As Variant
    Dim oMap As Variant
    oMap = CreateMap()

    ' ==== Основні дані ====
    Call MapPut(oMap, "CurrentDateField", "поточна дата і час")
    Call MapPut(oMap, "OffsetField",      "зсув")
    Call MapPut(oMap, "ReasonField",      "причина зсуву")
    Call MapPut(oMap, "DurationField",    "кількість днів")

    ' ==== Персональні дані ====
    Call MapPut(oMap, "LastNameField",    "прізвище")
    Call MapPut(oMap, "FirstNameField",   "ім'я")
    Call MapPut(oMap, "PatronymicField",  "по батькові")

    ' ==== Фінансові дані ====
    Call MapPut(oMap, "CodeCombo",        "код")
    Call MapPut(oMap, "PaidField",        "сплачено")
    Call MapPut(oMap, "ExpenseField",     "видаток")
    Call MapPut(oMap, "IncomeField",      "прихід")

    ' ==== Інше ====
    Call MapPut(oMap, "CommentField",     "коментар")
    Call MapPut(oMap, "PhoneField",       "телефон")
    Call MapPut(oMap, "PassportField",    "паспортні дані")
    Call MapPut(oMap, "BirthDateField",   "дата народження")

    ' ==== Службові поля таблиці (не у формі) ====
    Call MapPut(oMap, "CheckInDate",      "заселення")
    Call MapPut(oMap, "CheckOutDate",     "виселення")
    Call MapPut(oMap, "CreatedField",     "створено")
    Call MapPut(oMap, "BlacklistField",   "чорний список")
    Call MapPut(oMap, "HostelField",      "хостел")
    Call MapPut(oMap, "HistoryField",     "історія")

    GetFieldToColumnMap = oMap
End Function




