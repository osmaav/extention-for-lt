// ==UserScript==
// @name         Download Button for LT
// @namespace    http://tampermonkey.net
// @version      2025-06-12_v.3.7.1
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx
// @  Версия 3.7.1
// @  - Производительность: оптимизировать использование наблюдателей и сокращать количество операций над DOM.
// @  - Безопасность: контролировать ресурсные утечки и обеспечить совместимость с будущими изменениями сайта.
// @  - Интерфейс: улучшать взаимодействия пользователя и информативность уведомлений.
// @  - Кодовая база: уменьшить дублирование кода и упростить структуру CSS-стилей.
// @author       osmaav
// @homepageURL  https://github.com/osmaav/extention-for-lt
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @supportURL   https://github.com/osmaav/extention-for-lt/issues
// @match        https://www.leadertask.ru/web/*
// @match        https://www.beta.leadertask.ru/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(async () => {
  'use strict';

  // Подключение библиотеки XLSX через создание тега <script>
  function loadLibrary(callback) {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    script.onload = callback;
    document.head.appendChild(script);
  }

  // Глобально устанавливаем XLSX после успешного импорта
  loadLibrary(() => {
    window.XLSX = window.XLSX || {};
  });

  // Кэширование текущих значений для сравнения
  let oldCheckListSize = 0;
  let previousUrlPath = '';

  // Настройка CSS для кнопки
  function addStyles() {
    const styles = `
      .btnExpListToXlsx {
        background-color: rgba(0, 255, 0, 0.2);
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
        background-color: rgba(0, 255, 0, 0.1);
      }

      .btnExpListToXlsx:active {
        transform: scale(0.95);
      }
    `;

    const styleElem = document.createElement('style');
    styleElem.type = 'text/css';
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // Функционал кнопки
  function createDownloadButton() {
    const button = document.createElement('button');
    button.classList.add('btnExpListToXlsx');
    button.textContent = 'Скачать';
    button.onclick = handleDownloadClick;
    return button;
  }

  // Генерирует имя файла
  function generateFilename(taskName) {
    const dateStr = new Date().toLocaleDateString();
    return `CheckList-from-${taskName}-${dateStr}.xlsx`
      .replaceAll(',', '-')
      .replaceAll(':', '.');
  }

  // Экспорт чек-листа в Excel
  function exportToXlsx(taskName) {
    const rows = Array.from(getCheckList()).map((el, idx) => ({
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

  // Сбор чек-листа
  function getCheckList() {
    return document.querySelectorAll('#task-prop-content div.flex.items-center.w-full.group');
  }

  // Управление видимостью кнопки
  function manageButtonVisibility(checkListSize) {
    const button = document.querySelector('.btnExpListToXlsx');
    if (checkListSize > 2) {
      if (!button) {
        let targetEl = document.querySelector('#modal-container > div:nth-child(5) > div.flex > div > div > div:nth-child(2) > div > div > div:nth-child(2) > div:nth-child(2) > span');
        if (targetEl) {
          targetEl.append(createDownloadButton());
        }
        else {
          targetEl = document.querySelector('#modal-container #task-prop-content > div:nth-child(3) > div > span');
          if (targetEl) {
            targetEl.append(createDownloadButton());
          } else {
            document.querySelector('#modal-container #task-prop-content > div:nth-child(4) div span').append(createDownloadButton());
          }
        }
      }
    } else {
      if (button) {
        button.remove();
      }
    }
  }

  // Контролирует поведение кнопки при изменениях
  function handleDownloadClick(event) {
    event.preventDefault();
    const taskContainer = document.querySelector('.user_child_customer_custom div>div');
    const taskName = taskContainer.outerText
    .replaceAll(': ', '_')
    .replaceAll('/', '_')
    .replaceAll(' ', '_');
    exportToXlsx(taskName);
  }

  // Основной обработчик изменений
  function processChanges(urlChanged, checkListSizeChanged) {
    if (urlChanged || checkListSizeChanged) {
      manageButtonVisibility(getCheckList().length);
    }
  }

  // Настройка наблюдателя
  function setupMutationObserver() {
    const observerConfig = { attributes: true, childList: true, subtree: true };

    const modalObserver = new MutationObserver((mutations) => {
      let urlChanged = false;
      let checkListSizeChanged = false;

      mutations.forEach((mutation) => {
        if (mutation.type === 'childList') {
          const checkList = getCheckList();
          const currentCheckListSize = checkList.length;
          if (currentCheckListSize !== oldCheckListSize) {
            checkListSizeChanged = true;
            oldCheckListSize = currentCheckListSize;
          }
        }
      });

      const currentUrlPath = location.pathname;
      if (previousUrlPath !== currentUrlPath) {
        urlChanged = true;
        previousUrlPath = currentUrlPath;
      }

      processChanges(urlChanged, checkListSizeChanged);
    });

    modalObserver.observe(document.body, observerConfig);

    // Освобождение ресурса при выходе со страницы
    window.addEventListener('beforeunload', () => {
      modalObserver.disconnect();
    });
  }

  // Основной поток инициализации
  function init() {
    addStyles();
    setupMutationObserver();
  }

  // Запускаем основной процесс
  if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init(); // Если документ уже загружен, запускаем сразу
  } else {
    document.addEventListener('DOMContentLoaded', init); // Иначе ждем завершения загрузки
  }
})();
