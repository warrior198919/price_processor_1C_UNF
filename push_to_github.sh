#!/bin/bash
# Скрипт для автоматизации публикации проекта на GitHub

# Переход в директорию скрипта (корень проекта)
cd "$(dirname "$0")"

# Инициализация git, если не инициализирован
if [ ! -d ".git" ]; then
  git init
fi

# Добавление всех файлов
git add .

# Коммит с сообщением (если есть изменения)
if ! git diff --cached --quiet; then
  git commit -m "Initial commit"
fi

# Установка основной ветки
git branch -M main

# Добавление remote, если не добавлен
if ! git remote | grep -q origin; then
  git remote add origin "https://github.com/warrior198919/price_processor_1C_UNF.git"
fi

# Пуш на GitHub
# Если ветка уже существует на сервере, используем --force-with-lease для первого пуша
if git ls-remote --exit-code origin main &>/dev/null; then
  git push --force-with-lease origin main
else
  git push -u origin main
fi
