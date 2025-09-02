# Visarsoft Message Wizard - Installation Guide

## Що потрібно для встановлення

1. **Microsoft Outlook** (2016 або новіший)
2. **Веб-сервер** для хостингу файлів add-in'а
3. **OpenAI API ключ**

## Варіанти розгортання

### Варіант 1: Використання GitHub Pages (Рекомендується)

1. Завантажте папку `package` на ваш GitHub репозиторій
2. Увімкніть GitHub Pages в налаштуваннях репозиторію
3. Оновіть URL в `manifest.xml` на ваш GitHub Pages URL

### Варіант 2: Власний веб-сервер

1. Скопіюйте вміст папки `package` на ваш веб-сервер
2. Переконайтеся, що файли доступні через HTTPS
3. Оновіть URL в `manifest.xml` відповідно

## Інсталяція в Outlook

### Для розробників/тестування:

1. Відкрийте Outlook
2. Перейдіть до **Insert** → **Add-ins** → **Get Add-ins**
3. Виберіть **My add-ins** → **Add a custom add-in** → **Add from file**
4. Виберіть файл `manifest.xml` з папки `package`
5. Підтвердіть встановлення

### Для організацій:

1. Адміністратор має завантажити `manifest.xml` в Microsoft 365 Admin Center
2. Перейти до **Settings** → **Integrated apps** → **Upload custom apps**
3. Завантажити файл та розгорнути для потрібних користувачів

## Налаштування API ключа

Оскільки add-in використовує OpenAI API, потрібно налаштувати ключ:

1. Отримайте API ключ на https://platform.openai.com/account/api-keys
2. У файлах add-in'а замініть `VITE_OPENAI_API_KEY` на ваш ключ
3. **Увага**: В продакшені рекомендується використовувати backend сервер для API викликів

## Структура файлів для розгортання

```
package/
├── index.html          # Головний файл add-in'а
├── manifest.xml        # Маніфест для Outlook
└── assets/
    ├── index-xxx.css   # Стилі
    └── index-xxx.js    # JavaScript код
```

## Відлагодження

1. **Add-in не завантажується**: Перевірте що всі URL в manifest.xml правильні та доступні
2. **Помилка API ключа**: Переконайтеся що OpenAI ключ валідний та має достатню квоту
3. **Помилки CORS**: Переконайтеся що сервер налаштований для CORS запитів

## Підтримка

Для питань та підтримки зверніться до команди Visarsoft.