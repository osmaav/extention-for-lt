// ==UserScript==
// @name         Check-List->xlsx for LT
// @namespace    http://tampermonkey.net/
// @version      2025-06-12_v.3.7.0
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx
// Версия 3.7.0 
// - Производительность: оптимизировать использование наблюдателей и сокращать количество операций над DOM.
// - Безопасность: контролировать ресурсные утечки и обеспечить совместимость с будущими изменениями сайта.
// - Интерфейс: улучшать взаимодействия пользователя и информативность уведомлений.
// - Кодовая база: уменьшить дублирование кода и упростить структуру CSS-стилей.
// 
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

  // Импортируем библиотеку XLSX асинхронно
  try {
    const { XLSX } = await import('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
    window.XLSX = XLSX; // Сохраняем ссылку на XLSX в глобальном пространстве
  } catch (error) {
    console.error('Ошибка загрузки библиотеки XLSX:', error);
    return;
  }

  // Функция добавления CSS-кода
  function addStyles() {
    const styles = `
      .btnExpListToXlsx {
        background-color: var(--gray-200);
        border-radius: 6px;
        padding: 4px 8px;
        font-size: 14px;
        line-height: 16px;
        transition: all 0.3s0, ease-in-out;
        position: relative;
        margin-left: 5px;
        height: 1.6rem;
        width: 4.6rem;
      }
      
      .btnExpListToXlsx:hover {
        background-color: var(--gray-300);
      }
      
      .btnExpListToXlsx:active {
        transform: scale(0.95);
        box-shadow: inset 0 0 10px rgba(0, 0, 0, 0.2);
      }
    `;

    const styleElem = document.createElement('style');
    styleElem.type = 'text/css';
    styleElem.appendChild(document.createTextNode(styles));
    document.head.appendChild(styleElem);
  }

  // Создаем кнопку для скачивания
  function createDownloadButton() {
    const button = document.createElement('button');
    button.classList.add('btnExpListToXlsx');
    button.textContent = 'Скачать';
    button.onclick = downloadHandler;
    return button;
  }

  // Генерация имени файла
  function generateFilename(taskName) {
    const dateStr = new Date().toLocaleDateString();
    return `CheckList-from-${taskName}-${dateStr}.xlsx`
      .replaceAll(',', '-')
      .replaceAll(':', '.');
  }

  // Функция экспортирует данные в Excel
  function exportToXlsx(taskName) {
    const rows = Array.from(getCheckList()).map((el, idx) =>
      ({ idx: idx + 1, content: el.textContent.replace(/^[\d]+\.\s*/, '') })
    );

    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, taskName.slice(0, 31));
    XLSX.utils.sheet_add_aoa(sheet, [['Номер', 'Значение']], { origin: 'A1' });

    const filename = generateFilename(taskName);
    XLSX.writeFile(book, filename, { compression: true });
  }

  // Получаем чек-лист
  function getCheckList() {
    return document.querySelectorAll('#task-prop-content div.flex.items-center.w-full.group');
  }

  // Обработчик кликов по кнопке
  function downloadHandler(event) {
    event.preventDefault();
    const taskContainer = document.querySelector('.user_child_customer_custom div>div');
    const taskName = taskContainer.outerText
      .replaceAll(': ', '_')
      .replaceAll('/', '_')
      .replaceAll(' ', '_');
    exportToXlsx(taskName);
  }

  // Создание и отображение кнопки при наличии чек-листа
  function manageButtonVisibility() {
    const checkListItems = getCheckList();
    const hasCheckList = checkListItems.length > 2;

    if (hasCheckList) {
      const button = createDownloadButton();
      const targetEl = document.querySelector('#task-prop-content > div:nth-child(3) > div > span');
      if (targetEl) {
        targetEl.append(button);
      } else {
        document.querySelector('#task-prop-content > div:nth-child(4) > div > span').append(button);
      }
    } else {
      const existingButton = document.querySelector('.btnExpListToXlsx');
      if (existingButton) {
        existingButton.remove();
      }
    }
  }

  // Наблюдатель за изменением DOM
  function setupMutationObserver() {
    const config = { attributes: true, childList: true, subtree: true };
    const observer = new MutationObserver(manageButtonVisibility);
    observer.observe(document.body, config);
  }

  // Основной поток инициализации
  function init() {
    addStyles();
    setupMutationObserver();
    manageButtonVisibility(); // Первая проверка видимости кнопки
  }

  // Запускаем основной процесс
  if (document.readyState === 'complete') {
    init();
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
