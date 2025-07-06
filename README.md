# libreoffice-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
![Repo size](https://img.shields.io/github/repo-size/yourhostel/libreoffice-tools)
![Last commit](https://img.shields.io/github/last-commit/yourhostel/libreoffice-tools)

**LibreOffice macros for UI forms and automation.**

Набір макросів для LibreOffice для автоматизації заповнення таблиць через діалогові вікна.   
Логіка написана на LibreOffice Basic.

## Features
- Dialog-based UI creation
- Macros in LibreOffice Basic

## Structure
```txt
libreoffice-tools
├── basic
│   └── Standard
│       ├── Buttons.bas
│       │   └── AddButton
│       │
│       ├── CreateRecord.bas
│       │   ├── ShowForm
│       │   ├── InsertButton_actionPerformed
│       │   ├── InsertButton_disposing
│       │   ├── ValidateSelection
│       │   ├── OffsetReasonValidation
│       │   ├── OffsetReasonInsertion
│       │   ├── DateRangeInsertion
│       │   ├── PersonDataValidation
│       │   ├── PersonDataInsertion
│       │   ├── PaidInsertion
│       │   ├── FinanceAreNumbersValidation
│       │   ├── FinanceCommentValidation
│       │   ├── FinanceInsertion
│       │   ├── IsPhoneMinimalValid
│       │   ├── PhoneValidation
│       │   ├── PhoneInsertion
│       │   ├── PassportDataValidation
│       │   ├── BirthDateValidation
│       │   ├── DateFormatValidation
│       │   ├── PassportBirthInsertion
│       │   ├── HostelInsertion
│       │   ├── CodeInsertion
│       │   ├── PlaceInsertion
│       │   ├── CreateDialog
│       │   └── ShowPassportInvalid
│       │
│       ├── EditRecord.bas
│       │   ├── StartEdit
│       │   ├── IsInAllowedColumns
│       │   ├── IsInAllowedRows
│       │   ├── FindEditRange
│       │   └── MoveCursorToColumnA
│       │
│       ├── Events.bas
│       │   ├── AddTextFieldsOffsetListener
│       │   ├── OffsetField_textChanged
│       │   ├── OffsetField_disposing
│       │   ├── AddCodeComboListeners
│       │   ├── CodeCombo_mousePressed
│       │   ├── CodeCombo_mouseReleased
│       │   ├── CodeCombo_mouseEntered
│       │   ├── CodeCombo_mouseExited
│       │   ├── CodeCombo_itemStateChanged
│       │   ├── CodeCombo_disposing
│       │   ├── AddDurationComboListeners
│       │   ├── DurationCombo_mousePressed
│       │   ├── DurationCombo_mouseReleased
│       │   ├── DurationCombo_mouseEntered
│       │   ├── DurationCombo_mouseExited
│       │   ├── DurationCombo_itemStateChanged
│       │   ├── DurationCombo_disposing
│       │   ├── AddPlaceComboListeners
│       │   ├── PlaceCombo_mousePressed
│       │   ├── PlaceCombo_mouseReleased
│       │   ├── PlaceCombo_mouseEntered
│       │   ├── PlaceCombo_mouseExited
│       │   ├── PlaceCombo_itemStateChanged
│       │   └── PlaceCombo_disposing
│       │
│       ├── Fields.bas
│       │   ├── FieldTemplate
│       │   └── ComboBoxTemplate
│       │
│       ├── Filters.bas
│       │   ├── PeopleTodayFilter
│       │   ├── PersonWord
│       │   ├── CountVisibleRows
│       │   ├── ResetPeopleTodayFilter
│       │   ├── ResetFilter
│       │   └── GetPeopleRange
│       │
│       ├── Images.bas
│       │   └── AddLogo
│       │
│       ├── MainVariable.bas
│       │   └── GetFieldToColumnMap
│       │
│       ├── Notification.bas
│       │   └── ShowDialog
│       │
│       └── Utilities.bas
│           ├── Capitalize
│           ├── LockFields
│           ├── CreateMap
│           ├── MapPut
│           ├── MapGet
│           ├── MapHasKey
│           ├── AppendArray
│           ├── CalculatePaidFieldWithPlace
│           ├── UpdatePlaceCombo
│           ├── SelectFirstEmptyInA
│           ├── IsPlaceOccupiedToday
│           ├── DebugGun
│           └── CheckOccupiedPlace
│
├── .gitignore
├── README.md
└── LICENSE
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