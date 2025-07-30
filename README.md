# libreoffice-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
![Repo size](https://img.shields.io/github/repo-size/yourhostel/libreoffice-tools)
![Last commit](https://img.shields.io/github/last-commit/yourhostel/libreoffice-tools)

**LibreOffice macros for UI forms and automation.**

![2025-07-29 00-05-39.png](screenshots/2025-07-29%2000-05-39.png)

![2025-07-29 00-06-18.png](screenshots/2025-07-29%2000-06-18.png)

![2025-07-29 00-07-06.png](screenshots/2025-07-29%2000-07-06.png)

![2025-07-30 12-52-35.png](screenshots/2025-07-30%2012-52-35.png)

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