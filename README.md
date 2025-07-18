# libreoffice-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
![Repo size](https://img.shields.io/github/repo-size/yourhostel/libreoffice-tools)
![Last commit](https://img.shields.io/github/last-commit/yourhostel/libreoffice-tools)

**LibreOffice macros for UI forms and automation.**

## Structure
```txt
Standard
├── Blacklist
│   ├── BlacklistStart
│   ├── ShowBlacklistInstructions
│   ├── AddToBlacklist
│   ├── RemoveFromBlacklist
│   ├── FilterBlacklist
│   └── ToggleBlacklistFilter
│
├── CreateRecord
│   ├── StartCreate
│   ├── ShowForm
│   ├── InsertButton_actionPerformed
│   ├── InsertButton_disposing
│   ├── OffsetReasonValidation
│   ├── OffsetReasonInsertion
│   ├── DateRangeInsertion
│   ├── PersonDataValidation
│   ├── PersonDataInsertion
│   ├── PaidInsertion
│   ├── FinanceAreNumbersValidation
│   ├── FinanceCommentValidation
│   ├── FinanceInsertion
│   ├── IsPhoneMinimalValid
│   ├── PhoneValidation
│   ├── PhoneInsertion
│   ├── PassportDataValidation
│   ├── ShowPassportInvalid
│   ├── BirthDateValidation
│   ├── DateFormatValidation
│   ├── PassportBirthInsertion
│   ├── HostelInsertion
│   ├── CodeInsertion
│   ├── PlaceInsertion
│   ├── CreateDialog
│   ├── FormInitialization
│   └── ChoiceButtonName
│
├── DeleteRecord
│   └── DeleteRow
│
├── EditRecord
│   ├── StartEdit
│   ├── ReadFromTable
│   ├── ShowFieldsInMsgBox
│   ├── IsInAllowedColumns
│   ├── IsInAllowedRows
│   ├── isBelowEncashRange
│   ├── MoveCursorToColumnA
│   ├── FormatHistoryLine
│   └── AppendHistory
│
├── Encashment
│   ├── DoEncashment
│   ├── CalculateEncashment
│   ├── InsertEncashment
│   ├── GetAfterLastEncashRange
│   ├── AddFinancialRow
│   ├── ShowFinDialog
│   ├── ValidationFinData
│   └── FormatFinLine
│
├── Events
│   ├── AddTextFieldsOffsetListener
│   ├── OffsetField_textChanged
│   ├── OffsetField_disposing
│   ├── AddCodeComboListeners
│   ├── CodeCombo_mousePressed
│   ├── CodeCombo_mouseReleased
│   ├── CodeCombo_mouseEntered
│   ├── CodeCombo_mouseExited
│   ├── CodeCombo_itemStateChanged
│   ├── CodeCombo_disposing
│   ├── AddDurationComboListeners
│   ├── DurationCombo_mousePressed
│   ├── DurationCombo_mouseReleased
│   ├── DurationCombo_mouseEntered
│   ├── DurationCombo_mouseExited
│   ├── DurationCombo_itemStateChanged
│   ├── DurationCombo_disposing
│   ├── AddPlaceComboListeners
│   ├── PlaceCombo_mousePressed
│   ├── PlaceCombo_mouseReleased
│   ├── PlaceCombo_mouseEntered
│   ├── PlaceCombo_mouseExited
│   ├── PlaceCombo_itemStateChanged
│   └── PlaceCombo_disposing
│
├── Fields
│   ├── FieldTemplate
│   ├── ComboBoxTemplate
│   ├── AddButton
│   └── OptionGroupTemplate
│
├── Filters
│   ├── PeopleTodayFilter
│   ├── ShowCountResults
│   ├── PersonWord
│   ├── CountVisibleRows
│   ├── ResetFilter
│   ├── ResetPeopleTodayFilter
│   ├── GetRecordsRange
│   └── DebugRangeValues
│
├── Images
│   └── AddLogo
│
├── MainVariable
│   └── GetFieldToColumnMap
│
├── Notification
│   ├── ShowDialog
│   └── MsgDlg
│
└── Utilities
    ├── Capitalize
    ├── LockFields
    ├── CreateMap
    ├── MapPut
    ├── MapGet
    ├── MapHasKey
    ├── MapGetByIndex
    ├── AppendArray
    ├── MapClear
    ├── CalculatePaidFieldWithPlace
    ├── UpdatePlaceCombo
    ├── SelectFirstEmptyInA
    ├── FilterPlace
    ├── FilterCompetitors
    ├── CheckOccupiedPlace
    ├── ExcludeRow
    ├── ShowFields
    ├── DayWord
    ├── ShowPasswordDialog
    ├── GetOccupiedRows
    ├── GetVacantPlacesString
    ├── SetNextId
    ├── DiffArrays
    └── ShowArray
```

```bash
mkdir -p basic/Standard && \
touch basic/Standard/{Fields.bas,CreateRecord.bas,Buttons.bas,Images.bas}
```

![Знімок екрана з 2025-07-04 18-33-49.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-33-49.png)

![Знімок екрана з 2025-07-04 18-33-55.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-33-55.png)

![Знімок екрана з 2025-07-04 18-34-05.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-34-05.png)

![Знімок екрана з 2025-07-04 18-34-10.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-34-10.png)

![Знімок екрана з 2025-07-04 18-34-51.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-34-51.png)

![Знімок екрана з 2025-07-04 18-34-55.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-34-55.png)

![Знімок екрана з 2025-07-04 18-35-00.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-35-00.png)

![Знімок екрана з 2025-07-04 18-35-32.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-07-04%2018-35-32.png)

# Python

```bash
# Перевстановити з пайтон провайдером якщо потрібно
Перевстановити з пайтон провайдером якщо потрібно
sudo snap remove libreoffice
sudo apt remove --purge libreoffice*
sudo add-apt-repository ppa:libreoffice/ppa
sudo apt update
sudo apt install libreoffice libreoffice-script-provider-python

# Перевірити пайтон провайдер
dpkg -l | grep libreoffice-script-provider-python

# Якщо потрібні локалі
# Український інтерфейс
sudo apt install libreoffice-l10n-uk
# Російський інтерфейс
sudo apt install libreoffice-l10n-ru
# Усі локалі
sudo apt install libreoffice-l10n-*
```