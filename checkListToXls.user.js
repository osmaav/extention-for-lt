// ==UserScript==
// @name         Check-List->xlsx for LT v3.5.4 (2025-05-19)
// @namespace    http://tampermonkey.net/
// @version      3.5.4
// @description  Скрипт создает кнопку "скачать" для выгрузки Чек-листа в файл формата xlsx (версия 3.5.4 изменения: добавил логи для отладки событий)
// @author       osmaav
// @homepageURL  https://github.com/osmaav/extention-for-lt
// @updateURL    https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @downloadURL  https://raw.githubusercontent.com/osmaav/extention-for-lt/main/checkListToXls.user.js
// @supportURL   https://github.com/osmaav/extention-for-lt/issues
// @match        https://www.leadertask.ru/web/*
// @grant        none
// @run-at       document-idle

// ==/UserScript==
( async() => {
  'use strict';
  //console.warn('UserScript: Скрипт запущен');
  try {
    const { XLSX } = await import('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
  } catch (error) {
    console.warn('UserScript: ошибка загрузки модуля XLSX', error);
    return;
  }

  let btnExpListToXlsx = document.createElement('button');
  btnExpListToXlsx.classList.add('dark:bg-[#404040]',
                                 'border-[#1B1B1C0D]',
                                 'bg-[#F5F5F5]',
                                 'text-[14px]',
                                 'leading-[16px]',
                                 'py-[4px]',
                                 'px-[8px]',
                                 'dark:border-[#FFFFFF0D]',
                                 'border-[1px]',
                                 'border-solid',
                                 'rounded-[6px]',
                                 'hover:bg-[#E8E8E8]',
                                 'hover:dark:bg-[#333333]',
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

  function getcheckList() {
    //console.warn('UserScript: getcheckList is runnin', new Date().toLocaleString());
    try {
      return document.querySelectorAll('#task-prop-content div.flex.items-center.w-full.group');
    } catch (error) {console.warn('UserScript:', error);}
  }

  function exportToXlsx(taskName = 'Задача', xlsx ) {
    //console.warn('UserScript: exportToXlsx вызвана');
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

  function updateBtn(checkListLen) {
    //console.warn('UserScript: checkListLen', checkListLen);
    if (checkListLen > 2) { // -- чек-лист > 2
      if (document.querySelector('.btnExpListToXlsx')) { // -- если кнопка есть
        btnExpListToXlsx.style.display = 'block'; // -- показываем кнопку
        console.warn('UserScript: показали кнопку');
      } else { // -- если кнопки нет
        let targetEl = document.querySelector('#task-prop-content > div:nth-child(3) > div > span');// -- ищем целевой элемент к которому добавим кнопку
        console.warn('UserScript: targetEl найден', targetEl);
        if (targetEl) targetEl.append(btnExpListToXlsx);
        else {document.querySelector('#task-prop-content > div:nth-child(4) > div > span').append(btnExpListToXlsx);}
        // -- добавляем кнопку
        console.warn('UserScript: добавили кнопку');
      }
    } else if (checkListLen < 3) { // чек-лист < 3
      btnExpListToXlsx.style.display = 'none';// -- скрываем кнопку
      //console.warn('UserScript: скрыли кнопку');
    }
  }

  function MyMutationObserver() {
    console.warn('UserScript: DOMContentLoaded');
    let oldHref = '';
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
    //console.warn('UserScript: скрипт запущен');
    const bodyElement = document.querySelector('body');
    if (!bodyElement) {
      console.warn('UserScript: Элемент body не найден в DOM');
      return;
    }
    new MutationObserver(() => {
      console.warn('UserScript: новых событий поступило', mut.length, new Date().toLocaleString());
      let curHref = document.location.href;
      if (oldHref !== curHref) {
        console.warn('UserScript: путь изменился', curHref);
        oldHref = curHref;
        const taskPropertyWidow = document.querySelector(`#modal-container >div:nth-child(3)`);
        if (!taskPropertyWidow) {
          console.warn('UserScript: Элемент taskPropertyWidow не найден в DOM', new Date().toLocaleString());
          return;
        }
        if (!(taskPropertyWidow instanceof Node)) {
          console.warn('UserScript: Элемент taskPropertyWidow не является Node', new Date().toLocaleString());
          return;
        }
        new MutationObserver(mutations => {
          console.warn('UserScript: событий с измененным путем поступило', mutations.length, new Date().toLocaleString());
          for (const mutation of mutations) {
            let curHref = window.location.href;
            if (curHref.includes('/project/') || curHref.includes('/tasks/')) { // -- путь содержит project или tasks
              if (taskPropertyWidow.style?.display != 'none') { // -- окно открыто
                console.warn('UserScript: окно свойств открыто attributeName:', mutation.attributeName, 'type:',mutation.type, 'target:',mutation.target );
                if (mutation.attributeName === 'style') { // -- окно открылось
                  console.warn('UserScript: окно открылось', new Date().toLocaleString());
                  updateBtn(getcheckList().length);
                }// -- окно открылось
                if (mutation.type === 'childList') {
                  const removeLen = mutation.removedNodes[0]?.textContent.length;
                  //console.warn('UserScript: событие childList','target:',mutation.target);
                  let checkListLen = getcheckList().length;
                  if ((mutation.target.id === 'addNewCheckListEdit' || removeLen) && (checkListLen != oldCheckListLen)) { // -- чек-лист изменился
                    //console.warn('UserScript: Чек-лист изменился oldLen:', oldCheckListLen, 'newLen:', checkListLen);
                    updateBtn(checkListLen);
                    oldCheckListLen = checkListLen;
                  } // -- чек-лист изменился
                } // -- mutation.type === 'childList'
              } // -- окно открыто
            } // -- если путь содержит project или tasks
          }// -- обработка мутации
        }).observe(taskPropertyWidow, {
          attributes: true,
          subtree: true,
          attributeFilter: ['style'],
          childList: true
        });
      }
    }).observe(bodyElement, { childList: true, subtree: true });
  }
  if (document.readyState == 'loading') {
    // ещё загружается, ждём события
    document.addEventListener('DOMContentLoaded',MyMutationObserver());
  } else {
    // DOM готов!
    MyMutationObserver();
  }
})();
