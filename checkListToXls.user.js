// ==UserScript==
// @name         Download Button for LT
// @namespace    http://tampermonkey.net
// @version      2025-06-13_v.3.7.2
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx
// @  Версия 3.7.2
// @  - Производительность: оптимизировать использование наблюдателей и сокращать количество операций над DOM.
// @  - Вычисления: скорректировал получение списка для исключения дубликатов.
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
  function loadLibrary() {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    //script.onload = callback;
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
    // Получаем все подходящие элементы
    const elements = document.querySelectorAll('#task-prop-content div.flex.items-center.w-full.group');

    // Определяем индекс середины (делим длину массива пополам)
    const halfIndex = Math.floor(elements.length / 2);

    // Получаем первую половину элементов
    return [...elements].slice(0, halfIndex);
  }

  // Управление видимостью кнопки
  function manageButtonVisibility() {
    const button = document.querySelector('.btnExpListToXlsx');
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

  // Настройка наблюдателя
  // Функция для отслеживания изменений в DOM дереве
  function setupMutationObserver() {
    const modalContainer = document.querySelector('#modal-container');

    if (!modalContainer) {
      console.error('UserScript: Контейнер #modal-container не найден');
      return;
    }

    let prevLocation = location.pathname; // Храним начальную позицию URL

    const observer = new MutationObserver((mutations) => {
      // Работаем только с последним событием (последним элементом массива)
      const lastMutation = mutations[mutations.length - 1];

      if (lastMutation.type === 'childList') {
        // Находим третьего и пятого ребенка заново после изменений
        const thirdChild = modalContainer.children[2]; // :nth-child(3)
        const fifthChild = modalContainer.children[4]; // :nth-child(5)
        let windowOpen = false;

        if (thirdChild) {
          const display = window.getComputedStyle(thirdChild).display;
          if (display !== 'none') {
            console.log(`UserScript: Окно [3] открыто: ${display}`);
            windowOpen = true;
          }
        }

        if (fifthChild) {
          const display = window.getComputedStyle(fifthChild).display;
          if (display !== 'none') {
            console.log(`UserScript: Окно [5] открыто : ${display}`);
            windowOpen = true;

          }
        }
        if (windowOpen) {
          // Получаем текущий путь currentUrl
          const currentUrlPath = location.pathname;
          if (previousUrlPath !== currentUrlPath) {
            console.log('UserScript: измененился URL', currentUrlPath);
            previousUrlPath = currentUrlPath;
          }
          //показываю кнопку
          manageButtonVisibility();
        }
      }
    });

    // Следим за изменениями внутри #modal-container
    observer.observe(modalContainer, { childList: true, subtree: true, attributeFilter: ['style'] });
    //console.log('UserScript: Наблюдатель запущен');

//     // Периодическая проверка URL
//     const intervalId = setInterval(() => {
//       //console.log('UserScript: таймер запущен', intervalId);
//       const currentLocation = location.pathname;
//       if (prevLocation !== currentLocation) {
//         console.log('UserScript: измененился URL', currentLocation);
//         prevLocation = currentLocation;
//         console.log('UserScript: Наблюдатель отключен из-за изменения URL');
//         observer.disconnect(); // Отключение наблюдателя
//         // Следим за изменениями внутри #modal-container
//         observer.observe(modalContainer, { childList: true, subtree: true, attributeFilter: ['style'] });
//       }
//     }, 1000); // Проверять URL каждые полсекунды

    // Освобождение ресурса при выходе со страницы
    window.addEventListener('beforeunload', () => {
      console.log('UserScript: Наблюдатель отключен');
      observer.disconnect(); // Отключение наблюдателя
      //clearInterval(intervalId); // Остановка интервала
    });
  }



  // Основной поток инициализации
  function init() {
    // Настройка CSS для кнопки
    addStyles();
    // Запускаем наблюдатель
    setupMutationObserver();
  }

  // Запускаем основной процесс
  if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init(); // Если документ уже загружен, запускаем сразу
  } else {
    document.addEventListener('DOMContentLoaded', init); // Иначе ждем завершения загрузки
  }
})();
