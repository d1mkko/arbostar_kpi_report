Руководство по использованию скрипта отчета KPI Jira
1. Основные настройки и извлечение данных (JQL)
Основной файл это jira_kpi_report.py. В нем содержатся все настройки для подключения к Jira, определения команд и их участников, а также логика для извлечения данных.

1.1. Настройки подключения к Jira
В верхней части файла jira_kpi_report.py вы найдете секцию "Jira connection details":

# Jira connection details
JIRA_SERVER = 'https://arbostar.atlassian.net'
JIRA_EMAIL = '<your_email>'
JIRA_API_TOKEN = '<your_api_token' # Убедитесь, что здесь указан ваш актуальный API токен

JIRA_SERVER: URL вашего экземпляра Jira.

JIRA_EMAIL: Электронная почта, связанная с вашей учетной записью Jira.

JIRA_API_TOKEN: Ваш персональный токен API для аутентификации. Крайне важно, чтобы этот токен был актуальным и имел необходимые разрешения для доступа к данным в Jira.
Для генерации API Token - https://id.atlassian.com/manage-profile/security/api-tokens

1.2. Определение команд и участников
Далее вы увидите секцию TEAMS, где определены команды и их участники.

# Team members
# LDT TEAM
LDT_TEAM_MEMBERS = ['Andrew Belousov', 'Ivan Stepaniuk', 'serg levch']
# TWA TEAM
TWA_TEAM_MEMBERS = ['Anton Rozonenko', 'Anton Shelekhvost', 'Anton Shynkarenko', 
                    'Dmytro Yurchenko', 'Ivan Borovets', 'Maksim Levchenko', 
                    'Michael Parandiy', 'Oleg Lats', 'Oleg Nekrasov', 
                    'Oleksii Petrov', 'Roman Dubovka', 'Zubkov Pavlo']
# CWT TEAM
CWT_TEAM_MEMBERS = ['Sergey Chernov']
# BA TEAM
BA_TEAM_MEMBERS = ['Bohdan Kucher', 'Polina Reminna', 'Stepan Zhukevych']
# AMA TEAM
AMA_TEAM_MEMBERS = ['Andriy Momot', 'Arthur Hlushko', 'Denys Honchar', 
                'Iliya Sozonenko', 'Oleg Nekrasov', 'Oleksandr Korneiko', 
                'Oleksii Petrov']

# All teams and their members
TEAMS = {
    'LDT TEAM': LDT_TEAM_MEMBERS,
    'TWA TEAM': TWA_TEAM_MEMBERS,
    'CWT TEAM': CWT_TEAM_MEMBERS,
    'BA TEAM': BA_TEAM_MEMBERS,
    'AMA TEAM': AMA_TEAM_MEMBERS
}

Убедитесь, что списки _TEAM_MEMBERS содержат точные имена участников Jira (отображаемые имена).

Словарь TEAMS связывает имена команд с соответствующими списками участников.

1.3. JQL-запросы для извлечения данных
Основная логика извлечения данных из Jira определяется в словаре TASK_CATEGORIES. Здесь для каждой категории задач (например, "ASAP Changes", "BugFixes") прописаны специфические JQL-запросы.

Основные JQL-запроса хранятся в словаре TASK_CATEGORIES. Пример:

TASK_CATEGORIES = {
    'ASAP Changes': {
        'query': '(project = "TWA" OR project = "LDT" OR project = "CWT") AND issuetype="Change request" AND cf[ВАШ_ID_ПОЛЯ_РЕЛИЗ]=ASAP  and statusCategoryChangedDate >= "-21d" and statusCategoryChangedDate <= "-7d"  and  status not in ("DEV", "Merge to Staging", "Staging", "HOTFIX", "Merge to Master", "MASTER", "Ready for release") and assignee = {assignee}',
        'ama_query': '(project = "AMA") AND issuetype="Change request" and cf[ВАШ_ID_ПОЛЯ_РЕЛИЗ]=ASAP and statusCategoryChangedDate >= "-21d" and statusCategoryChangedDate <= "-7d" and  status not in ("Ready to test", "Test passed", "Test pre-release", "Ready for release", "Released", "Cancelled") and assignee = {assignee}'
    },
    # ... другие категории ...
}

Что здесь важно:

cf[ВАШ_ID_ПОЛЯ_РЕЛИЗ] и cf[ВАШ_ID_ПОЛЯ_ЭПИК_ЛИНК]: Это ОЧЕНЬ ВАЖНЫЕ места для корректировки! cf[YOUR_RELEASE_FIELD_ID] и cf[YOUR_EPIC_LINK_FIELD_ID] должны быть заменены на ЧИСЛОВЫЕ ID ваших пользовательских полей "Release" (или "Dropdown Release", "Target Release" и т.п.) и "Epic Link" в вашей Jira. Без этого JQL-запросы будут выдавать ошибки.

Как найти ID пользовательского поля в Jira: Зайдите в "Настройки Jira" (шестеренка вверху справа) -> "Задачи" -> "Пользовательские поля". Найдите нужное поле, нажмите на три точки (...) и выберите "Контексты и значения по умолчанию" или "Изменить". В URL страницы будет числовой ID поля (например, customFieldId=10102).

project = "TWA" OR project = "LDT" OR project = "CWT": Определяет проекты, из которых будут извлекаться задачи для данной категории. Убедитесь, что ключи проектов (например, "TWA", "LDT") соответствуют вашим проектам в Jira.

issuetype="...": Фильтр по типу задачи (например, "Change request", "Bug", "Task").

statusCategoryChangedDate >= "-21d" and statusCategoryChangedDate <= "-7d": Определяет временной диапазон для отчета (предыдущий спринт: от 21 дня назад до 7 дней назад). Эти значения (PREV_SPRINT_START и PREV_SPRINT_END) задаются глобально в скрипте.

status not in (...): Исключает задачи с определенными статусами.

assignee = {assignee}: Это плейсхолдер, который скрипт автоматически заменяет на имя каждого члена команды при выполнении JQL.

1.4. Извлечение данных о затраченном времени (Worklogs)
Скрипт также собирает данные о затраченном времени (worklogs) для каждого члена команды. Это делается в функции get_tracked_time_for_period.

def get_tracked_time_for_period(jira, date_start_relative, date_end_relative, team_members):
    # ...
    all_projects = set()
    for team_name_key in TEAMS.keys():
        for category in TEAM_CATEGORIES.get(team_name_key, []):
            category_info = TASK_CATEGORIES.get(category)
            if category_info:
                import re
                project_matches_query = re.findall(r'project\s*=\s*"([^"]+)"', category_info.get('query', ''))
                project_matches_ama_query = re.findall(r'project\s*=\s*"([^"]+)"', category_info.get('ama_query', ''))
                all_projects.update(project_matches_query)
                all_projects.update(project_matches_ama_query)
                
    project_jql_clause = ""
    if all_projects:
        project_jql_clause_content = " OR ".join([f'project = "{p}"' for p in all_projects])
        project_jql_clause = f"({project_jql_clause_content}) AND "
    # ...
    jql_broad_issues = (
        f"{project_jql_clause} worklogDate >= '{date_start_relative}' AND worklogDate <= '{date_end_relative}' "
        f"AND worklogAuthor in ({assignee_list_for_jql})"
    )
    # ...
    issues_to_check = jira.search_issues(jql_broad_issues, maxResults=False, fields='summary,worklog,assignee')
    # ...

Автоматическое определение проектов: Скрипт автоматически собирает все уникальные ключи проектов из ваших TASK_CATEGORIES и использует их для формирования project_jql_clause. Это гарантирует, что запросы worklog будут ограничены только теми проектами, которые вы отслеживаете в своих категориях задач.

worklogDate: Фильтрует записи о времени по дате их создания.

worklogAuthor: Фильтрует записи по автору (члену команды).

fields='summary,worklog,assignee': Запрашивает у Jira конкретно поле worklog, которое содержит информацию о затраченном времени. Затем скрипт проходит по каждой записи worklog для каждой задачи, суммируя timeSpentSeconds и конвертируя их в часы.

2. Как выполнить скрипт
Для упрощения выполнения скрипта предусмотрен bash-файл run_kpi_report.sh. Он автоматизирует установку зависимостей и последовательный запуск обоих Python-скриптов.

2.1. Подготовка перед запуском
Убедитесь, что у вас установлен Python 3.

Сохраните все файлы (jira_kpi_report.py, jira_kpi_report_pie_gen.py, run_kpi_report.sh, requirements.txt) в одну и ту же папку.

ОБЯЗАТЕЛЬНО отредактируйте jira_kpi_report.py и замените плейсхолдеры. Убедитесь, что JIRA_SERVER, JIRA_EMAIL и JIRA_API_TOKEN корректны.

Убедитесь, что флаги использования моковых данных в jira_kpi_report.py установлены в False для получения реальных данных из Jira:

USE_MOCK_BA_DATA = False
USE_MOCK_AMA_DATA = False
USE_MOCK_OTHER_DATA = False

2.2. Пошаговое выполнение
Откройте терминал (или командную строку в Windows).

Перейдите в папку, где сохранены все файлы скрипта, используя команду cd:

cd /путь/к/вашей/папке/со/скриптами

(Например: cd /Users/ваше_имя/Documents/JiraReport)

Выполните bash-файл:

bash run_kpi_report.sh

(На Windows вам может понадобиться установить WSL (Windows Subsystem for Linux) или запускать Python-скрипты по отдельности: python jira_kpi_report.py затем python jira_kpi_report_pie_gen.py).

2.3. Что делает run_kpi_report.sh?
Файл run_kpi_report.sh выполняет следующие действия:

cd "$(dirname "$0")": Переходит в директорию, где находится сам скрипт.

Проверка Python 3:

if ! command -v python3 &> /dev/null
then
    echo "Python 3 not found. Please install Python 3."
    exit 1
fi

Проверяет наличие Python 3 в вашей системе. Если не найден, выдает ошибку и завершает работу.

Установка зависимостей:

echo "📦 Installing Python dependencies..."
python3 -m pip install -r requirements.txt

Устанавливает все необходимые Python-библиотеки (такие как jira, pandas, openpyxl), перечисленные в файле requirements.txt. Это нужно сделать только один раз или при изменении зависимостей.

Запуск jira_kpi_report.py:

echo "🔄 Generating sprint report..."
python3 jira_kpi_report.py

Запускает основной скрипт, который подключается к Jira, извлекает данные согласно JQL-запросам, обрабатывает их и создает файл sprint_report.xlsx. В процессе выполнения вы увидите логи, показывающие выполнение JQL-запросов и количество найденных задач/worklogs.

Проверка создания отчета и запуск jira_kpi_report_pie_gen.py:

REPORT_FILE="sprint_report.xlsx"
if [ -f "$REPORT_FILE" ]; then
    echo "✅ Report created: $REPORT_FILE"
    echo "📊 Adding pie charts..."
    python3 jira_kpi_report_pie_gen.py
    # ...
else
    echo "❌ Error: Report file not found. Pie charts were not added."
fi

Проверяет, был ли создан файл sprint_report.xlsx. Если да, то запускает jira_kpi_report_pie_gen.py, который открывает сгенерированный отчет и добавляет в него круговые диаграммы, представляющие вклад каждой категории задач в общий объем по статусам.

Автоматическое открытие отчета (только macOS):

if [[ "$OSTYPE" == "darwin"* ]]; then
    open "$REPORT_FILE"
fi

Если вы используете macOS, скрипт попытается автоматически открыть сгенерированный Excel-файл.

После успешного выполнения скрипт создаст файл sprint_report.xlsx в той же директории, откуда вы его запустили.

Структура выходного файла sprint_report.xlsx
Отчет будет содержать несколько листов:

Summary: Основной лист с агрегированными данными по каждой команде, показывающий количество задач по статусам ("To Do", "In Development", "Completed", "Declined", "Cancelled"), Story Points и затраченное время для каждого члена команды. Также включает сводную таблицу статусов по всем командам.

Task Details: Подробный лист со списком всех задач, извлеченных скриптом, их ключами, названиями, статусами, исполнителями, Story Points и предполагаемым временем.

ChartData: Скрытый лист, используемый для генерации данных для круговых диаграмм.
