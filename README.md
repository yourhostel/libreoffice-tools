# libreoffice-tools

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
![Repo size](https://img.shields.io/github/repo-size/yourhostel/libreoffice-tools)
![Last commit](https://img.shields.io/github/last-commit/yourhostel/libreoffice-tools)

**LibreOffice macros for UI forms and automation.**

![2025-07-29 00-05-39.png](screenshots/2025-07-29%2000-05-39.png)

![2025-07-29 00-06-18.png](screenshots/2025-07-29%2000-06-18.png)

![2025-07-29 00-07-06.png](screenshots/2025-07-29%2000-07-06.png)

![2025-07-30 12-58-40.png](screenshots/2025-07-30%2012-58-40.png)

`deploy.sh` Скрипт розпаковує таблицю, що містить локальні модулі і поміщає ці модулі для використання глобально в стандартну папку $HOME/.config/libreoffice/4/user/basic/Standard. Дана процедура дозволяє правильно перетворити написані текстові файли, які зазвичай подаються як .bas, у формат зрозумілий libreoffice - формат .xba. Можна написати парсер, але треба враховувати крім валідної шапки XML символи, що потребують екранування:

| Символ | Назва          | Екранована форма |
| ------ | -------------- | ---------------- |
| `'`    | апостроф       | `&apos;`         |
| `"`    | подвійна лапка | `&quot;`         |
| `&`    | амперсанд      | `&amp;`          |
| `<`    | менше          | `&lt;`           |
| `>`    | більше         | `&gt;`           |

#### Приклад:
- `.bas`
```
' =====================================================
' === Функція IsCancelCode =============================
' =====================================================
Function IsCancelCode(nCode As Long) As Boolean
    IsCancelCode = InStr(" 20 21 22 23 ", " " & nCode & " ") > 0
End Function
```

- `.xba`
```
&apos; =====================================================
&apos; === Функція IsCancelCode =============================
&apos; =====================================================
Function IsCancelCode(nCode As Long) As Boolean
    IsCancelCode = InStr(&quot; 20 21 22 23 &quot;, &quot; &quot; &amp; nCode &amp; &quot; &quot;) &gt; 0
End Function
```
- Для встановлення потрібно у файлі `deploy.sh` визначити змінну `ODS`, посилання на таблицю контейнера модулів. Виконати скрипт:
```bash
chmod +x deploy.sh
./deploy.sh
```

# Python
```bash
# Перевстановити з пайтон провайдером якщо потрібно
Перевстановити з пайтон провайдером якщо потрібно
sudo snap remove libreoffice
sudo apt rem[screenshots](screenshots)ove --purge libreoffice*
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