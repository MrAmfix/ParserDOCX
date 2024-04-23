Schemas.py - там хранятся схемы для работы с docx
Например пишешь:
from schemas import schemas

И для итерации параграфов например:
for para in root.iter(f'{{{schemas.w}}}p'):
Или можно использовать так же как словарь (schemas['w'])

Так же я добавил 2 папки в gitignore:
documents_for_extract (туда распакуй все файлы для парсинга)
out_json (там будут все JSON-файлы, можешь по коду посмотреть)
