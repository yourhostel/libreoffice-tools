REM  *****  BASIC  *****

' MainVariable.bas

Public Const HOSTEL = "Саперка"
Public Const NEGET = 123

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