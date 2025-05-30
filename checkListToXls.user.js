// ==UserScript==
// @name         Check-List->xlsx for LT
// @namespace    http://tampermonkey.net/
// @version      2025-05-30_v.3.6.4
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx (версия 3.6.4 - исправил адрес для поддержки beta версии)
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
( async() => {
  'use strict';
  //console.warn('UserScript: Скрипт запущен', currtime());
  try {
    const { XLSX } = await import('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
  } catch (error) {
    console.warn('UserScript:',currtime(), 'ошибка загрузки модуля XLSX', error);
    return;
  }

  let btnExpListToXlsx = document.createElement('button');
  btnExpListToXlsx.classList.add('dark:bg-[#1B1B1C]/[0.5]',
                                 'dark:hover:bg-[#1B1B1C]/[0.8]',
                                 'bg-gray-200',
                                 'hover:bg-gray-300',
                                 'text-[14px]',
                                 'leading-[16px]',
                                 'py-[4px]',
                                 'px-[8px]',
                                 'border-solid',
                                 'rounded-[6px]',
                                 'btnExpListToXlsx');
  btnExpListToXlsx.style.marginLeft = '5px';
  btnExpListToXlsx.style.position = 'relative';//
  btnExpListToXlsx.style.height = '1.6rem';//
  btnExpListToXlsx.style.width = '4.6rem';//
  btnExpListToXlsx.innerHTML = 'Скачать';

  let myDiv = document.createElement('div');
  myDiv.classList.add('el-popper', 'is-dark');
  myDiv.dataset.popperPlacement = 'bottom';
  myDiv.style.display = 'none';
  myDiv.style.left = '-3.8rem';//
  myDiv.innerHTML = 'Скачать список в виде таблицы';

  let mySpan = document.createElement('span');
  mySpan.classList.add('el-popper__arrow');

  myDiv.appendChild(mySpan);
  btnExpListToXlsx.appendChild(myDiv);
  btnExpListToXlsx.addEventListener('click',
    event => {
      event.stopPropagation();
      exportToXlsx(document
                   .querySelector('.user_child_customer_custom div>div')
                   .outerText.replaceAll(': ', '_')
                   .replaceAll('/', '_')
                   .replaceAll(' ', '_')
                  , XLSX);
    },
  true);
  let myTable = [];
  let flMyStyleAdd = false;
  let oldCheckListLen = 0;
  function currtime() {
      const now = new Date();
      const currt = now.toLocaleString() + '.' + now.getMilliseconds();
      return currt};
  function getcheckList() {
    //console.warn('UserScript:',currtime(), 'getcheckList is runnin', currtime());
    try {
      return document.querySelectorAll('#task-prop-content div.flex.items-center.w-full.group');
    } catch (error) {console.warn('UserScript: ',currtime(), error);}
  }

  function exportToXlsx(taskName = 'Задача', xlsx ) {
    //console.warn('UserScript:',currtime(), 'exportToXlsx вызвана');
    Array.from(getcheckList())
      .map(el => el.textContent)
      .map((el , idx) => myTable.push({ idx: idx + 1, content: el.replace(new RegExp('/^[0-9]+.s+/', 'g'), '') }))
      .filter(el => el != undefined);
    const worksheet = xlsx.utils.json_to_sheet(myTable);
    const workbook = xlsx.utils.book_new();
    //console.log(`Список готов!`, myTable);
    xlsx.utils.book_append_sheet(workbook, worksheet, taskName.slice(0, 31));
    xlsx.utils.sheet_add_aoa(worksheet, [['Номер', 'Значение']], { origin: 'A1' });
    const max_width_col_A = myTable.reduce((w, r) => Math.max(w, r?.idx.toString().length), 5);
    const max_width_col_B = myTable.reduce((w, r) => Math.max(w, r?.content.length), 30);
    myTable = [];
    worksheet['!cols'] = [{ wch: max_width_col_A }, { wch: max_width_col_B }];
    let fileName = `CheckList-from-${taskName}-${new Date().toLocaleString()}.xlsx`
    .replaceAll(new RegExp(/,\s+/, 'g'), '-')
    .replaceAll(new RegExp(/\:/, 'g'), '.');
    xlsx.writeFile(workbook, fileName, {
      compression: true
    });
  }


  function updateBtn() {
    let checkListLen = getcheckList().length;
    //console.warn('UserScript:', currtime(), 'updateBtn is running...');
    if (checkListLen > 2) { // -- чек-лист > 2
      if (document.querySelector('.btnExpListToXlsx')) { // -- если кнопка есть
        btnExpListToXlsx.style.display = 'block'; // -- показываем кнопку
        console.warn('UserScript:',currtime(), 'показали кнопку');
      } else { // -- если кнопки нет
        let targetEl = document.querySelector('#task-prop-content > div:nth-child(3) > div > span');// -- ищем целевой элемент к которому добавим кнопку
        //console.warn('UserScript: targetEl найден', targetEl);
        if (targetEl) targetEl.append(btnExpListToXlsx);
        else {document.querySelector('#task-prop-content > div:nth-child(4) > div > span').append(btnExpListToXlsx);}
        // -- добавляем кнопку
        console.warn('UserScript:',currtime(), 'добавили кнопку');
      }
    } else if (checkListLen < 3) { // чек-лист < 3
      btnExpListToXlsx.style.display = 'none';// -- скрываем кнопку
      console.warn('UserScript:', currtime(), 'скрыли кнопку');
    }
  }

  function MyMutationObserver() {
    console.warn('UserScript:',currtime(), 'DOMContentLoaded');
    let oldUrl = '';
    const css = `
      .btnExpListToXlsx>div {
         top: 125%;
         left: -63%;
         overflow: inherit;
         text-wrap-mode: nowrap;
         text-align: center;
       }
      .btnExpListToXlsx:hover>div {
         display: flex !important;
         animation: opacity 1.5s infinite;
         animation-direction: alternate;
       }
       @keyframes opacity{
        from{
          color: #707173;
        }
        to {
          color: white;
        }
       }
       .btnExpListToXlsx>div>span {
         left: 5.5rem;
       }
       .btnExpListToXlsx:hover>div>span {
         animation-direction: alternate;
       }
       @keyframes opacity2{
        to {
          opacity: 1;
        }
       }`;
    let style = document.createElement('style');
    style.type = 'text/css';
    style.appendChild(document.createTextNode(css));
    document.head.appendChild (style);
    const app = document.querySelector('#app');

    if (!app) {
      console.warn('UserScript:', currtime(), 'Элемент app не найден в DOM');
      return;
    }

    function throttle(fn, delay = 1000) {
      //console.warn('UserScript:', currtime(), 'throttle is running...');
      let lastCallTime = 0;
      return function(...args) {
        const currentTime = Date.now();
        if (currentTime - lastCallTime >= delay) {
          console.warn('UserScript:', currtime(), 'fn is running...', fn, args);
          fn.apply(this, args);
          lastCallTime = currentTime;
        }
      };
    }

    new MutationObserver(() => {
      let curUrl = document.location.href;
//       const taskPropertyWidow = document.querySelector(`#modal-container`);
      const taskPropertyWidow = document.querySelector(`#modal-container`);
      if (oldUrl != curUrl) {
//        console.warn('UserScript:',currtime(), 'путь изменился', oldUrl, curUrl);
        oldUrl = curUrl;
        if (!taskPropertyWidow) {
          console.warn('UserScript:', currtime(), 'taskPropertyWidow не найден', taskPropertyWidow);
          return;
        }
        if (!(taskPropertyWidow instanceof Node)) {
          console.warn('UserScript:',currtime(), 'taskPropertyWidow не является Node');
          return;
        }
        //if (!taskPropertyWidow.style.length) console.warn('UserScript:',currtime(), 'обрабатываем события:');
        const throttledUpdateBtn = throttle(updateBtn, 1000);

        new MutationObserver(mutations => {
//           console.warn('UserScript:', currtime(), 'обрабатываем события:', mutations);
          // -- путь содержит project или tasks
          if (curUrl.includes('/project/') || curUrl.includes('/tasks/')) {
            //console.warn('UserScript:', currtime(), 'путь содержит project или tasks', curUrl);
            // -- обработка событий
            for (const mutation of mutations) {
              if ((mutation.attributeName === 'style') && (mutation.type === 'attributes')) { // -- окно открылось
                if ((curUrl.includes('/project/')) && (curUrl.split('/').length === 7)) {
                  throttledUpdateBtn();
                } // -- путь содержит project
                if ((curUrl.includes('/tasks/')) && (curUrl.split('/').length === 6)) {
                  throttledUpdateBtn();
                }// -- путь содержит tasks
              }// -- окно открылось
              if (mutation.type === 'childList') {
                const removeLen = mutation.removedNodes[0]?.textContent.length;
                //console.warn('UserScript:', currtime(), 'событие childList','target:',mutation.target);
                let checkListLen = getcheckList().length;
                if ((mutation.target.id === 'addNewCheckListEdit' || removeLen) && (checkListLen != oldCheckListLen)) { // -- чек-лист изменился
                  //console.warn('UserScript:', currtime(), 'Чек-лист изменился oldLen:', oldCheckListLen, 'newLen:', checkListLen);
                  throttledUpdateBtn();
                } // -- чек-лист изменился
              } // -- mutation.type === 'childList'
            }// -- обработка событий
          } // -- если путь содержит project или tasks
        }).observe(taskPropertyWidow, {
          attributes: true,
          subtree: true,
          attributeFilter: ['style'],
          childList: true
        });

      }
    }).observe(app, {subtree: true, attributeFilter: ['style'], childList: true});
  }
  if (document.readyState == 'loading') {
    // ещё загружается, ждём события
    document.addEventListener('DOMContentLoaded', MyMutationObserver());
  } else {
    // DOM готов!
    MyMutationObserver();
  }
})();
