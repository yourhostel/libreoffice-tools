# libreoffice-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
![Repo size](https://img.shields.io/github/repo-size/tysser/libreoffice-tools)
![Last commit](https://img.shields.io/github/last-commit/tysser/libreoffice-tools)

**LibreOffice macros for UI forms and automation.**

## Features
- Dialog-based UI creation
- Macros in LibreOffice Basic

## Structure
- `AddLogo` — image control
- `AddButton` — buttons
- `FieldTemplate` — label + input field


```bash
mkdir -p basic/Standard && \
touch basic/Standard/{Fields.bas,Settlement.bas,Buttons.bas,Images.bas}
```

```txt
libreoffice-tools
├── basic
│   └── Standard
│       ├── Buttons.bas
│       ├── Fields.bas
│       ├── Images.bas
│       └── Settlement.bas
├── .gitignore
├── README.md
└── LICENSE
```

![Знімок екрана з 2025-06-27 13-07-00.png](screenshots/%D0%97%D0%BD%D1%96%D0%BC%D0%BE%D0%BA%20%D0%B5%D0%BA%D1%80%D0%B0%D0%BD%D0%B0%20%D0%B7%202025-06-27%2013-07-00.png)