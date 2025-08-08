// ==UserScript==
// @name         Download Button for LT 3.11
// @version      2025-08-08_v.3.11
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx
// @author       osmaav
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @match        https://beta.leadertask.ru/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=leadertask.ru
// @grant        none
// @run-at       document-idle
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
      /* Стили для кнопки */
      .btnExpListToXlsx {
        /* background-color: rgba(0, 255, 0, 0.2);*/
        border-radius: 6px;
        padding: 4px 8px;
        font-size: 14px;
        line-height: 16px;
        transition: all 0.3s ease-in-out;
        position: relative;
        margin-left: 5px;
        height: 1.6rem;
        width: 4.6rem;
        box-shadow: inset 0 0 3px rgba(0, 0, 0, 0.3);
      }

      .btnExpListToXlsx:hover {
        /*background-color: rgba(0, 255, 0, 0.1);*/
      }

      .btnExpListToXlsx:active {
        transform: scale(0.95);
      }
    `;

    const styleElem = document.createElement('style');
    // styleElem.type = 'text/css';
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // 4. Создание кнопки скачивания
  function createDownloadButton() {
    const button = document.createElement('button');
    button.classList.add('btnExpListToXlsx', 'dark:bg-[#0A0A0C]', 'dark:text-[#C5C6CF]', 'opacity-50', 'hover:opacity-100');
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
    const elements = document.querySelectorAll('#task-prop-content [contenteditable="true"]');
    return [...elements];
  }

  // 8. Управление видимостью кнопки
  function manageButtonVisibility() {
    const button = document.querySelector('.btnExpListToXlsx');
    if (!button) {
      // Сначала выбираем все подходящие элементы по уникальным классам
      const candidates = document.querySelectorAll('#modal-container #task-prop-content span.flex.w-full.border-solid');

      // Затем фильтруем их по наличию текста "Чек-лист"
      const targetEl = [...candidates].find(el => el.textContent.includes('Чек-лист'));

      if (targetEl) {
        targetEl.append(createDownloadButton());
      }
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
  function setupMutationObserver() {
    const modalContainer = document.querySelector('#modal-container'); // Найти контейнер с идентификатором "#modal-container"

    if (!modalContainer) { // Проверьте наличие контейнера
      console.error('UserScript: Контейнер #modal-container не найден');
      return;
    }

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
    setupMutationObserver();
  }

  // 12. Инициализация основного процесса
  if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
