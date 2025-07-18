REM  *****  BASIC  *****

' MainVariable.bas

Public Const HOSTEL = "Саперка"
Public Const ENCASH = "інкасація"
Public Const BALANCE = "баланс"
Public Const PATH_TO_LOGO = "file:///home/tysser/Documents/LibreOfficeAssets/logo_1.png"
Public Const VALID_DURATIONS = "1;2;3;4;5;6;7;14;21;28"
Public Const NEGET_RULES = 123
Public Const LIST_OF_CODES = "1;2;3;4;5;6;8;9;10;11"
Public Const EXCLUDED_CODES = "7" ' "7;20;30;8"
Public Const MAX_SEARCH_RANGE_IN_PRICE = 28
Public Const ACTION_CREATE = "create"
Public Const ACTION_EDIT = "edit"
Public Const ACTION_CHECK_ROW = "check_row"
Public Const ALL_PLACES = "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28"
Public Const ALL_EXPENSES = "зарплата;повернення;хімія;розхідники;зв’язок;запчастини;обладнання;ремонт;реклама;комуналка;оренда;охорона;податки;ліцензії;штрафи;страховка"
Public Const ALL_INCOMES  = "товари;оренда;послуги;депозити;харчування"
Public       PLACE_COMBO_HEIGHT As Long

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




