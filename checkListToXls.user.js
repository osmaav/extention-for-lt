// ==UserScript==
// @name         Download Button for LT 4.5.1
// @version      2025-12-12_v.4.5.1
// @description  Скрипт создает кнопку для выгрузки Чек-листа в файл формата xlsx
// @author       osmaav
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @match        https://*.beta.leadertask.ru/*
// @match        https://www.leadertask.ru/web/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=leadertask.ru
// @grant        none
// @run-at       document-idles
// ==/UserScript==

// Этот скрипт добавляет кнопку с иконкой на веб-страницу Leadertask, позволяющую экспортировать чек-лист в файл формата xlsx.
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
      .btnExpListToXlsx {
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
        top: -3px;
        border: var(--border-width, 1px) solid transparent;
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

      html.dark .btnExpListToXlsx::before{
          background: #111;
        }
      }
    `;

    const styleElem = document.createElement('style');
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // 4. Обработка кликов на кнопке
  function handleDownloadClick(event) {
    event.preventDefault();
    const taskContainer = document.querySelector('.user_child_customer_custom div>div');
    const taskName = taskContainer.outerText
      .replaceAll(': ', '_')
      .replaceAll('/', '_')
      .replaceAll(' ', '_');
    exportToXlsx(taskName, event.target.offsetParent);
  }
  
  // 5. Создание кнопки скачивания
  function createDownloadButton(target) {
    const button = document.createElement('button');
    const classes = ["btnExpListToXlsx", "bg-[#EEEEF1]", "dark:bg-[#0A0A0C]", "opacity-50", "hover:opacity-100"];
    classes.forEach(el=>button.classList.add(`${el}`));
    button.innerHTML = `<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAACXBIWXMAAAsTAAALEwEAmpwYAAABY0lEQVR4nN2Uv0oDQRDG09gpgmCjYmMj+hRaKb6AjQ9gaWNsBEEQn8PChxAxePd9k6ibGYWAtahlQloLiWy8YC45ktuYWDjwVfPnt7Mzu4XCvzPSdkB9o1hrsLQBsYNgAKivw4vbj2gnYR0kiXnjKNaC6PlEAQyBjApgG5JjJrkB1LgPQK2PDRCUl381bajGs5oSCOg4RapbED31isu63UlI7311l3SrpBWDARC7r1Selrwo9tALAPWzVKpNU+yd1ONggFdU0bU4tvXuBKZX8QJwKxC7zPg+osEdUBHHbtkLYuwHaIPUI4q9iNTmkpwPiB0C1YXM4ikAbIO0s289bva2nMxgz3eJsu4nf1GxMMx+s0VITt7umBpnAvzGjAbRKPiBTuQl/ykAOa4KYrc+VsQWQWv2+an1KNL5kecBsWsf65ybAvUuAwDvC72FTHPOzVKs3AUn8DwzluJpiN6QdtVb/At3PVyfwezqAwAAAABJRU5ErkJggg==" alt="xls-export">`;
    button.onclick = handleDownloadClick;
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
    const elements = parent.querySelectorAll('#task-prop-content [contenteditable][placeholder="Добавить"]');
    return [...elements];
  }
  
  // 8. Экспорт чек-листа в Excel
  function exportToXlsx(taskName, parent) {
    const checklist = getCheckList(parent);
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
