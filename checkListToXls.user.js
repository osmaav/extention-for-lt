// ==UserScript==
// @name         Download Button for LT 4.5.3
// @version      2026-04-01_v.4.5.3
// @description  Скрипт создает кнопку для скачивания Чек-листа в файл формата xlsx
// @author       osmaav
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @match        https://*.beta.leadertask.ru/*
// @match        https://www.leadertask.ru/web/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=leadertask.ru
// @grant        none
// @run-at       document-idles
// ==/UserScript==

// Этот скрипт добавляет кнопку с иконкой на веб-страницу Leadertask, позволяющую скачать чек-лист в файл формата xlsx.
// Скрипт подключается к библиотеке XLSX, обрабатывает события и контролирует отображение кнопки.

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
      .btnExpListToXlsx {
        user-select: none;
        -moz-user-select: none;
        -khtml-user-select: none;
        -webkit-user-select: none;
        -o-user-select: none;
        font-size: 0.9em;
        position: relative;
        z-index: 0;
        padding: 0.5em 0.5em;
        left: 0.5em;
        top: 0px;
        border-radius: 6px;
        transition: all 0.3s ease-in-out;
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
      }

      .btnExpListToXlsx::after{
        z-index: -2;
        inset: -1px;
      }

      .svgColor{
         color: #0F0;
      }
        
      .btnExpListToXlsx:hover{
        .svgColor{
         color: #FF4C00;
        }
      }
    `;

    const styleElem = document.createElement('style');
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // 4. Обработка кликов на кнопке
  function handleDownloadClick(target) {
    const taskContainer = document.querySelector('.user_child_customer_custom div>div');
    const taskName = taskContainer.outerText
      .replaceAll(': ', '_')
      .replaceAll('/', '_')
      .replaceAll(' ', '_');
    exportToXlsx(taskName, target.parentElement.parentElement.parentElement.parentElement);
  }
  
  // 5. Создание кнопки скачивания
  function createDownloadButton(target) {
    const button = document.createElement('button');
    button.classList.add(
        'btnExpListToXlsx',
        'bg-[#EEEEF1]',
        'dark:bg-[#0A0A0C]',
        'opacity-50',
        'hover:opacity-100',
    );

    // Отображаем изображение в качестве иконки загрузки
    const iconSVG = `
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" class="svgColor">
  <path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path>
  <path d="M7 10L12 15L17 10" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path>
  <path d="M12 15V3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path>
</svg>
    `;

    button.innerHTML = iconSVG;
    button.title = 'Скачать чек-лист'; // Альтернативный текст для доступности

    // Присваиваем обработчик нажатия кнопки
    button.onclick = handleDownloadClick.bind(null, target);
      
    return button;
  }
  
  // 6. Генерация имени файла
  function generateFilename(taskName) {
    const dateStr = new Date().toLocaleDateString();
    return `CheckList-from-${taskName}-${dateStr}.xlsx`
      .replaceAll(',', '-')
      .replaceAll(':', '.');
  }

  // 7. Получение элементов чек-листа
  function getCheckList(parent) {
    const elements = parent.querySelectorAll('#task-prop-content [contenteditable]');
    return [...elements];
  }
  
  // 8. Экспорт чек-листа в Excel
  function exportToXlsx(taskName, parent) {
    const checklist = getCheckList(parent);
    if (!checklist.length) return;
    const rows = Array.from(checklist).map((el, idx) => ({
      idx: idx+1,
      content: el.textContent.trim()
    }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, taskName.slice(0, 31));
    XLSX.utils.sheet_add_aoa(sheet, [['Номер', 'Значение']], { origin: 'A1' });
    const filename = generateFilename(taskName);
    XLSX.writeFile(book, filename, { compression: true });
  }

  // 9. Управление видимостью кнопки
  function manageButtonVisibility(target) {
    const button = target.querySelector('.btnExpListToXlsx');
    if (!button) {
      target.append(createDownloadButton(target));
    }
  }

  // 10. Установка наблюдателя за изменениями DOM
  function setupMutationObserver(target) {
    const observer = new MutationObserver((mutations) => { // Создать экземпляр наблюдателя
      const lastMutation = mutations[mutations.length - 1]; // Получить последнее событие изменения
      if (lastMutation.type === 'childList') { // Если произошло изменение списка дочерних элементов
        const targets = target.querySelectorAll('#task-prop-content span');
        if (targets.length) { // Если окно открыто
          const currentUrlPath = location.pathname; // Текущий путь URL
          if (previousUrlPath !== currentUrlPath) { // Проверить изменение пути
            previousUrlPath = currentUrlPath; // Обновляем предыдущее значение пути
            targets.forEach(el => {if (el.textContent.includes('Чек-лист')) manageButtonVisibility(el)}); // Показываем кнопку скачивания
          }
        }
      }
    });

    observer.observe(target, { childList: true, subtree: true, attributeFilter: ['style'] }); // Включить наблюдение за изменением детей и стилями

    window.addEventListener('beforeunload', () => { // Убедимся, что наблюдатель отключится при закрытии страницы
      observer.disconnect();
    });
  }

  // 11. Основная логика запуска
  function init() {
    addStyles();
    function findTarget(){
      const target = document.querySelector('#modal-container') // ждем появления контейнера для окна свойств
      if (target) {
        clearInterval(interval)
        setupMutationObserver(target)
      }
    }
    const interval = setInterval(findTarget, 500)
  }

  // 12. Инициализация основного процесса
  if (document.readyState === 'interactive' || document.readyState === 'complete') {
    init();
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
