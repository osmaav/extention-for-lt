// ==UserScript==
// @name         Download Button for LT 4.4.3
// @version      2025-11-18_v.4.4.3
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx
// @author       osmaav
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @match        https://*.beta.leadertask.ru/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=leadertask.ru
// @grant        none
// @run-at       document-idles
// ==/UserScript==

// Этот скрипт добавляет кнопку "Скачать" на веб-страницу Leadertask, позволяющую экспортировать чек-лист в файл формата xlsx.
// Скрипт подключается к библиотеке XLSX, обрабатывает события и контролирует отображение кнопки экспорта.

(async () => {
  'use strict';
  // 1. Подключение библиотеки XLSX
  function loadLibrary() {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    document.head.appendChild(script);
  }

  // Устанавливаем глобальное значение XLSX после успешной загрузки библиотеки
  loadLibrary(() => {
    window.XLSX = window.XLSX || {};
  });

  // 2. Начальные настройки
  let previousUrlPath = '';

  // 3. Добавление стилей для кнопки скачивания
  function addStyles() {
    const styles = `
    @property --color-1 {
        syntax: "<color>";
        inherits: true;
        initial-value: red;
      }

      @property --color-2 {
        syntax: "<color>";
        inherits: true;
        initial-value: yellow;
      }

      @property --color-3 {
        syntax: "<color>";
        inherits: true;
        initial-value: green;
      }

      @property --color-4 {
        syntax: "<color>";
        inherits: true;
        initial-value: blue;
      }

      @property --color-5 {
        syntax: "<color>";
        inherits: true;
        initial-value: purple;
      }

      @property --glow-deg {
        syntax: "<angle>";
        inherits: true;
        initial-value: 0deg;
      }

      @keyframes glow {
        100% {
          --glow-deg: 360deg;
        }
      }

      .btnExpListToXlsx {
        --gradient-glow:
          var(--color-1),
          var(--color-2),
          var(--color-3),
          var(--color-4),
          var(--color-5),
          var(--color-1);
        --glow-intensity: 0.125;
        --glow-size: 8px;
        --border-width: 1px;
        user-select: none;
        -moz-user-select: none;
        -khtml-user-select: none;
        -webkit-user-select: none;
        -o-user-select: none;
        font-size: 0.9em;
        position: relative;
        z-index: 0;
        padding: 0.1em 0.5em;
        left: 0.5em;
        top: 0px;
        border: var(--border-width, 1px) solid transparent;
        border-radius: 6px;
        background: linear-gradient(white, white) padding-box,
          conic-gradient(from var(--glow-deg), var(--gradient-glow)) border-box;
        transition: all 0.3s ease-in-out;
        animation: glow 10s infinite linear;
      }

      .btnExpListToXlsx::before,
      .btnExpListToXlsx::after{
        content: '';
        position: absolute;
        border-radius: inherit;
      }

      .btnExpListToXlsx::before{
        z-index: -1;
        inset: 1px;
        background: white;
        filter: blur(var(--glow-size, 6px));
      }

      .btnExpListToXlsx::after{
        z-index: -2;
        inset: -1px;
        background: conic-gradient(from var(--glow-deg), var(--gradient-glow));
        filter: blur(var(--glow-size, 6px));
        opacity: var(--glow-intensity, 0.5);
      }

      .btnExpListToXlsx:hover {
        --glow-intensity: 0.5;
        --glow-size: 2px;
      }

      .btnExpListToXlsx:hover:active {
        font-weight: bold;
        background: linear-gradient(to top, rgba(0, 0, 0, 0), rgba(0, 0, 0, 0)) padding-box,
        conic-gradient(from var(--glow-deg), var(--gradient-glow)) border-box;
      }

      html.dark
        .btnExpListToXlsx {
          background: linear-gradient(black, black) padding-box,
            conic-gradient(from var(--glow-deg), var(--gradient-glow)) border-box;
      }

      html.dark .btnExpListToXlsx::before{
          background: #111;
        }
      }
    `;

    const styleElem = document.createElement('style');
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // 4. Создание кнопки скачивания
  function createDownloadButton() {
    const button = document.createElement('button');
    button.classList.add('btnExpListToXlsx');
    button.textContent = 'Скачать';
    button.onclick = handleDownloadClick;
    return button;
  }

  // 5. Генерация имени файла
  function generateFilename(taskName) {
    const dateStr = new Date().toLocaleDateString();
    return `CheckList-from-${taskName}-${dateStr}.xlsx`
      .replaceAll(',', '-')
      .replaceAll(':', '.');
  }

  // 6. Экспорт чек-листа в Excel
  function exportToXlsx(taskName) {
    const checklist = getCheckList();
    if (!checklist.length) return;
    const rows = Array.from(checklist).map((el, idx) => ({
      idx: idx + 1,
      content: el.textContent.trim()
    }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, taskName.slice(0, 31));
    XLSX.utils.sheet_add_aoa(sheet, [['Номер', 'Значение']], { origin: 'A1' });
    const filename = generateFilename(taskName);
    XLSX.writeFile(book, filename, { compression: true });
  }

  // 7. Получение элементов чек-листа
  function getCheckList() {
    const elements = document.querySelectorAll('#task-prop-content [contenteditable][placeholder="Добавить"]');
    return [...elements];
  }

  // 8. Управление видимостью кнопки
  function manageButtonVisibility() {
    const button = document.querySelector('.btnExpListToXlsx');
    if (!button) {
      // Выбираем все подходящие элементы и фильтруем их по наличию текста "Чек-лист"
      document.querySelectorAll('#modal-container #task-prop-content span').forEach(el => {if (el.textContent.includes('Чек-лист')) el.append(createDownloadButton())});
    }
  }

  // 9. Обработка кликов на кнопке
  function handleDownloadClick(event) {
    event.preventDefault();
    const taskContainer = document.querySelector('.user_child_customer_custom div>div');
    const taskName = taskContainer.outerText
      .replaceAll(': ', '_')
      .replaceAll('/', '_')
      .replaceAll(' ', '_');
    exportToXlsx(taskName);
  }

  // 10. Установка наблюдателя за изменениями DOM
  function setupMutationObserver(modalContainer) {
    const observer = new MutationObserver((mutations) => { // Создать экземпляр наблюдателя
      const lastMutation = mutations[mutations.length - 1]; // Получить последнее событие изменения
      if (lastMutation.type === 'childList') { // Если произошло изменение списка дочерних элементов
        const thirdChild = modalContainer.children[2]; // Найти третий элемент (#modal-container > nth-child(3))
        const fifthChild = modalContainer.children[4]; // Найти пятый элемент (#modal-container > nth-child(5))
        let windowOpen = false; // Флаг для проверки открытия окон

        if (thirdChild && window.getComputedStyle(thirdChild).display !== 'none') { // Проверить видимость третьего элемента
          windowOpen = true; // Открытие подтверждено
        }

        if (fifthChild && window.getComputedStyle(fifthChild).display !== 'none') { // Проверить видимость пятого элемента
          windowOpen = true; // Открытие подтверждено
        }

        if (windowOpen) { // Если одно из окон открыто
          const currentUrlPath = location.pathname; // Текущий путь URL
          if (previousUrlPath !== currentUrlPath) { // Проверить изменение пути
            previousUrlPath = currentUrlPath; // Обновляем предыдущее значение пути
          }
          manageButtonVisibility(); // Показываем кнопку скачивания
        }
      }
    });

    observer.observe(modalContainer, { childList: true, subtree: true, attributeFilter: ['style'] }); // Включить наблюдение за изменением детей и стилями

    window.addEventListener('beforeunload', () => { // Убедимся, что наблюдатель отключится при закрытии страницы
      observer.disconnect();
    });
  }

  // 11. Основная логика запуска
  function init() {
    addStyles();
    let modalContainer;
    function findmodalContainer(){
      modalContainer = document.querySelector('#modal-container')
      if (modalContainer) {
        clearInterval(interval)
        setupMutationObserver(modalContainer)
      }
    }
    const interval = setInterval(findmodalContainer, 300)
  }

  // 12. Инициализация основного процесса
  if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
