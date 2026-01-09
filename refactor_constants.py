import re

file_path = 'd:/code/cytesi/measurement_analyzer_github/measurement_analyzer.py'

with open(file_path, 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

replacements = {
    'APP_TITLE': 'AppConfig.TITLE',
    'LOG_FILENAME': 'AppConfig.LOG_FILENAME',
    'THEME_CONFIG_FILE': 'AppConfig.THEME_CONFIG_FILE',
    'COL_FILE': 'AppConfig.Columns.FILE',
    'COL_TIME': 'AppConfig.Columns.TIME',
    'COL_NO': 'AppConfig.Columns.NO',
    'COL_PROJECT': 'AppConfig.Columns.PROJECT',
    'COL_MEASURED': 'AppConfig.Columns.MEASURED',
    'COL_DESIGN': 'AppConfig.Columns.DESIGN',
    'COL_DIFF': 'AppConfig.Columns.DIFF',
    'COL_UPPER': 'AppConfig.Columns.UPPER',
    'COL_LOWER': 'AppConfig.Columns.LOWER',
    'COL_RESULT': 'AppConfig.Columns.RESULT',
    'COL_UNIT': 'AppConfig.Columns.UNIT',
    'COL_ORIGINAL_JUDGE': 'AppConfig.Columns.ORIGINAL_JUDGE',
    'COL_ORIGINAL_JUDGE_PDF': 'AppConfig.Columns.ORIGINAL_JUDGE_PDF',
}

# Apply replacements only after line 80 to avoid touching AppConfig definition
start_line = 80
new_lines = lines[:start_line]
content_to_process = "".join(lines[start_line:])

for old, new in replacements.items():
    # Use regex to ensure we match whole words
    pattern = r'\b' + re.escape(old) + r'\b'
    content_to_process = re.sub(pattern, new, content_to_process)

new_lines.append(content_to_process)

with open(file_path, 'w', encoding='utf-8-sig') as f:
    f.writelines(new_lines)

print("Refactoring complete.")
