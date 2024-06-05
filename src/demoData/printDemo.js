import detailData from './detailData.js'

export const checkISExcelAreaBorder = function(x) {
  return (
    (x.borderType === 'border-outside' &&
      x.color === '#FF715E' &&
      x.rangeType === 'range' &&
      x.selfType === 'excel-area' &&
      x.style === '4') ||
    (x.borderType === 'border-outside' &&
      x.color === '#ff0000' &&
      x.rangeType === 'range' &&
      x.style === '4')
  );
};

// 获取打印区域内的变量，区域外变量
export const innerExcelAreaChange = (sheet) => {
  const borderInfo = sheet?.config?.borderInfo;

  let borderIndexArr = [];
  let setlistListVariables = [];
  let setlistPlainVariables = [];

  if (borderInfo) {
    const excelAreaBorderArr = borderInfo.filter(x => checkISExcelAreaBorder(x));

    if (excelAreaBorderArr.length > 0) {
      for (let i = 0; i < excelAreaBorderArr.length; i++) {
        const borderInfoItem = excelAreaBorderArr[i];
        const range = borderInfoItem.range?.[0];
        if (range) {
          // const selectCol_first = range?.column[0];
          // const selectCol_last = range?.column[1];
          // const selectRow_first = range?.row[0];
          // const selectRow_last = range?.row[1];
          const sheetData = sheet?.data;
          if (sheetData?.length > 0) {
            let startIndex = range?.row[0];
            let endIndex = Math.min(range?.row[1], sheetData.length);
            borderIndexArr.push({
              startIndex,
              endIndex,
            });
            // 获取在excel区域设置的变量
            for (let i = startIndex; i <= endIndex; i++) {
              const rowData = sheetData[i];
              if (rowData?.length > 0) {
                for (let j = 0; j < rowData.length; j++) {
                  const cell = rowData[j];
                  if (cell) {
                    if (cell?.custom) {
                      setlistListVariables.push(cell.enN);
                    } else if (cell?.m && cell.m.startsWith('{') && cell.m.endsWith('}')) {
                      const enName = cell.m.replace('{', '').replace('}', '');
                      setlistListVariables.push(enName);
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  const checkISInsideExcelAreaBorder = function(borderIndexArr, pos) {
    for (let i = 0; i < borderIndexArr.length; i++) {
      const borderItem = borderIndexArr[i];
      if (pos >= borderItem?.startIndex && pos <= borderItem?.endIndex) {
        return true;
      }
    }
    return false;
  };

  if (borderIndexArr?.length > 0) {
    const sheetData = sheet?.data;
    if (sheetData?.length > 0) {
      for (let i = 0; i <= sheetData?.length; i++) {
        if (checkISInsideExcelAreaBorder(borderIndexArr, i)) {
          continue;
        }

        const rowData = sheetData[i];
        if (rowData?.length > 0) {
          for (let j = 0; j < rowData.length; j++) {
            const cell = rowData[j];
            if (cell) {
              if (cell?.custom) {
                setlistPlainVariables.push(cell.enN);
              } else if (cell?.m && cell.m.startsWith('{') && cell.m.endsWith('}')) {
                const enName = cell.m.replace('{', '').replace('}', '');
                setlistPlainVariables.push(enName);
              }
            }
          }
        }
      }
    }
  } else {
    const sheetData = sheet?.data;
    if (sheetData?.length > 0) {
      for (let i = 0; i <= sheetData?.length; i++) {
        const rowData = sheetData[i];
        if (rowData?.length > 0) {
          for (let j = 0; j < rowData.length; j++) {
            const cell = rowData[j];
            if (cell) {
              if (cell?.custom) {
                setlistPlainVariables.push(cell.enN);
              } else if (cell?.m && cell.m.startsWith('{') && cell.m.endsWith('}')) {
                const enName = cell.m.replace('{', '').replace('}', '');
                setlistPlainVariables.push(enName);
              }
            }
          }
        }
      }
    }
  }

  return {
    setlistListVariables,
    setlistPlainVariables,
  };
};

const amountArr = ['unitPrice','taxAmount', 'amount', 'actualAmount', 'paidAmount', 'exempt', 'discount', 'lateFee', 'subLateFee', 'thisMonthFee', 'lastMonthFee', 'acLateFee', 'acUnPaid']

export const checkIsAmount = (name) => {
  return amountArr.includes(name)
}

let printHtmlRef = document.createDocumentFragment();

let lcRef = false;

let pageSizeRef = ''

const initLuckysheet = async (data) => {
  const luckysheet = window['luckysheet'];
  if (luckysheet && !lcRef) {
    lcRef = true;
    return new Promise(rev => {
      luckysheet.create({
        container: 'luckysheet',
        title: '打印模板',
        showinfobar: false,
        enableAddRow: false, // 允许添加行
        enableAddBackTop: false, // 允许添加返回顶部
        showsheetbar: false, // 是否展示底部sheet页
        defaultFontSize: 14,
        lang: 'zh',
        allowEdit: false, // 是否允许编辑
        // plugins:['chart'],
        // 默认显示1个sheet
        data: [data],
        showtoolbar: false, // 是否显示工具栏
      });

      rev(true);
    });
  } else {
    return new Promise(resolve => {
      setTimeout(() => {
        luckysheet.setSheetAdd({
          sheetObject: data,
          order: 0,
          success: () => {
            resolve(true);
          },
        });
      });
    });
  }
};

export const renderExcel = (detail = detailData) => {
  printHtmlRef = document.createDocumentFragment();
  if (detail && Array.isArray(detail)) {
    const render = async () => {
      const luckysheet = window['luckysheet'];
      console.time("2");
      for (let i = 0; i < detail.length; i++) {
        const item = detail[i];
        const OrginTemplateData = item?.templateVo?.templateData;
        const pageWidth = item?.templateVo?.printWidth;
        const pageHeight = item?.templateVo?.printHeight;

        if(pageWidth && pageHeight) {
          pageSizeRef = `
            @page {
              size: ${pageWidth / 10}mm ${pageHeight / 10}mm;
              margin: 0mm !important;
            }
          `
        } else {
          pageSizeRef = `
            @page {
              margin: 0mm !important;
            }
          `
        }

        if (OrginTemplateData) {
          const templateData = JSON.parse(OrginTemplateData);
          // console.log('templateData', templateData);
          const { setlistListVariables, setlistPlainVariables } = innerExcelAreaChange(
            templateData
          );

          // 删除设置的打印区域边框
          const sheet = JSON.parse(OrginTemplateData);
          if (sheet?.config?.borderInfo?.length > 0) {
            for (let i = 0; i < sheet?.config?.borderInfo.length; i++) {
              const item = sheet?.config?.borderInfo[i];
              if (checkISExcelAreaBorder(item)) {
                sheet?.config?.borderInfo.splice(i, 1);
                i--;
              }
            }
          }

          const sheetJson = JSON.stringify(sheet);

          const data = item?.data ?? [];

          const DataLen = data.length ?? 0;

          for (let j = 0; j < DataLen; j++) {

            const dataItem = data[j];

            await initLuckysheet(JSON.parse(sheetJson));

            // 替换打印区域外的变量
            for (let i = 0; i < templateData.celldata?.length; i++) {
              const row = templateData.celldata[i];
              const name = row.v.v;
              const target = setlistPlainVariables.find(x => {
                return (
                  x === row?.v?.enN || (name && x === name.replace('{', '').replace('}', ''))
                );
              });

              if (target) {
                const isAmount = checkIsAmount(target)
                const value = dataItem[target];
                if (value !== null || value !== undefined) {
                  luckysheet.setCellValue(row.r, row.c, 
                    {
                      m: isAmount ? '¥ ' + value : value,
                      v: value
                    }
                  , {isRefresh: false});
                } else {
                  luckysheet.setCellValue(row.r, row.c, '', {isRefresh: false});
                }
              }
            }

            /**
             * 替换打印区域内的变量-Start
             */
            const rowNumber = templateData.celldata.filter(s => {
              const name = s.v.v;
              return (
                setlistListVariables.indexOf(s?.v?.enN) > -1 ||
                (name &&
                  setlistListVariables.indexOf(name.replace('{', '').replace('}', '')) > -1)
              );
            })?.[0]?.r;

            const listLegth = dataItem?.dataList?.length || 0;

            const rowStart = rowNumber;
            const rowEnd = rowNumber + listLegth;

            const colArr = templateData.celldata
              .filter(s => {
                const name = s.v.v;
                return (
                  setlistListVariables.indexOf(s?.v?.enN) > -1 ||
                  (name &&
                    setlistListVariables.indexOf(name.replace('{', '').replace('}', '')) > -1)
                );
              })
              .map(c => c.c);

            // 获取列的起始点
            const startCol = Math.min(...colArr);
            const endCol = Math.max(...colArr);

            if (rowStart) {
              // 添加行
              for (let ii = 0; ii < listLegth - 1; ii++) {
                luckysheet.insertRow(rowStart);
              }
            }

            for (let jj = rowStart; jj < rowEnd; jj++) {
              for (let k = startCol; k <= endCol; k++) {
                const listVariable =
                  setlistListVariables[setlistListVariables.length - 1 - (endCol - k)];

                const isAmount = checkIsAmount(listVariable)

                const value = dataItem.dataList?.[listLegth - (rowEnd - jj)]?.[listVariable];

                if (value !== null || value !== undefined) {
                  luckysheet.setCellValue(jj, k, {
                    m: isAmount ? '¥ ' + value : value,
                    v: value
                  }, {isRefresh: jj === rowEnd - 1 && k === endCol});
                } else {
                  luckysheet.setCellValue(jj, k, '', {isRefresh: jj === rowEnd - 1 && k === endCol});
                }
              }
            }

            // 删掉占位行
            // luckysheet.deleteRow(rowStart + listLegth, rowStart + listLegth);
            /**
             * 替换打印区域内的变量-End
             */

            // 此方法慢
            // sheetData2Img(i * DataLen + j != 0);
            sheetData2HtmlDiv(i * DataLen + j != 0);
          }
        }
      }
      console.timeEnd("2");
      handlePrint()
    };
    render();
  }
}

const sheetData2Img = isDivide => {
  const luckysheet = window['luckysheet'];

  // const sheets = luckysheet.getAllSheets();
  const sheet = luckysheet.getSheet();

  const rowArr = sheet.celldata.map(c => c.r);
  const columnArr = sheet.celldata.map(c => c.c);

  const rowMin = Math.min(...rowArr);
  const rowMax = Math.max(...rowArr);

  const colMin = Math.min(...columnArr);
  const colMax = Math.max(...columnArr);


  const img = luckysheet.getScreenshot({
    range: { row: [0, rowMax], column: [colMin ?? 0, colMax] },
  }); // 需要传递参数 例如 A1:B1. 否则默认为选中区域

  const div = document.createElement('div');
  if (isDivide) {
    div.classList.add('page-container-last');
  }

  const el = document.createElement('img');
  el.src = img;
  // el.style.maxWidth = '100%';
  el.style.objectFit = 'contain';
  div.style.position = 'relative';

  // 移除不明边框线
  const bt = document.createElement('div');
  bt.style.position = 'absolute';
  bt.style.width = '100%';
  bt.style.background = '#ffffff';
  bt.style.height = '2px';
  bt.style.top = '-2px';
  bt.style.zIndex = '10';
  bt.style.padding = '2px';

  const bl = document.createElement('div');
  bl.style.position = 'absolute';
  bl.style.height = '100%';
  bl.style.background = '#ffffff';
  bl.style.width = '2px';
  bl.style.left = '-2px';
  bl.style.zIndex = '10';
  bl.style.padding = '2px';

  div.appendChild(bt);
  div.appendChild(bl);

  div.appendChild(el);

  // let extraImages:any = [];
  // if (sheet.images) {
  //   extraImages = Object.keys(sheet.images).map((s) => sheet.images[s]);

  //   extraImages.forEach((i) => {
  //     var iel = document.createElement('img');
  //     iel.src = i.src;
  //     iel.width = i.default.width;
  //     iel.height = i.default.height;

  //     var left = i.default.left;
  //     var top = i.default.top;

  //     iel.style.left = left + 'px';
  //     iel.style.top = top + 'px';
  //     iel.style.position = 'absolute';

  //     div.appendChild(iel);
  //   });
  // }

  div.style.marginBottom = 40 + 0 + 'px';

  printHtmlRef.appendChild(div);
};

const sheetData2HtmlDiv = isDivide => {
  const luckysheet = window['luckysheet'];
  const sheet = luckysheet.getSheet();

  const rowArr = sheet.celldata.map(c => c.r);
  const columnArr = sheet.celldata.map(c => c.c);

  const rowMin = Math.min(...rowArr);
  const rowMax = Math.max(...rowArr);

  const colMin = Math.min(...columnArr);
  const colMax = Math.max(...columnArr);

  const html = luckysheet.getRangeHtml({
    range: { row: [0, rowMax], column: [colMin ?? 0, colMax] },
  }); // 需要传递参数 例如 A1:B1. 否则默认为选中区域

  const div = document.createElement('div');
  if (isDivide) {
    div.classList.add('page-container-last');
  }
  div.style.position = 'relative';

  div.innerHTML = html;

  // let extraImages:any = [];
  // if (sheet.images) {
  //   extraImages = Object.keys(sheet.images).map((s) => sheet.images[s]);

  //   extraImages.forEach((i) => {
  //     var iel = document.createElement('img');
  //     iel.src = i.src;
  //     iel.width = i.default.width;
  //     iel.height = i.default.height;

  //     var left = i.default.left;
  //     var top = i.default.top;

  //     iel.style.left = left + 'px';
  //     iel.style.top = top + 'px';
  //     iel.style.position = 'absolute';

  //     div.appendChild(iel);
  //   });
  // }

  // div.style.marginBottom = 40 + 0 + 'px';

  printHtmlRef.appendChild(div);
};
  
const handlePrint = () => {
  var pel = document.createElement('div');
  pel.style.display = 'flex';
  pel.style.alignItems = 'center';
  pel.style.justifyContent = 'center';
  pel.style.flexDirection = 'column';
  pel.style.margin = '0';
  pel.style.padding = '0 20px';

  pel.appendChild(printHtmlRef);

  // 创建iframe 打印
  const printContent = pel.outerHTML;

  const printHTML =
    `<html><head><title>' '</title>` +
    `<style media="print">
      *{
        padding: 0;
        margin: 0;
        font-family: unset !important;
      }
      table {
        table-layout: fixed;
        border-collapse: collapse;
        max-width: 100%;
      }
      td {
        padding: 4px;
        max-width: 100px;
        white-space: nomarl;
        word-break: break-all;
        vertical-align: top !important;
        font-size: 10pt !important;
      }
      tr {
        page-break-inside: avoid;
        font-size: 10pt !important;
      }

      .page-container-last {
        page-break-before: always;
      }

      @media print {
        ${pageSizeRef}
        body {
          margin: 0;
        }
      }
    </style>
    ` +
    '</head><body>' +
    printContent +
    '</body></html>';
  const iframe = document.createElement('iframe');
  iframe.setAttribute('style', 'position: absolute; width: 0; height: 0;');
  document.body.appendChild(iframe);
  const iframeDoc = iframe.contentWindow.document;

  // 写入内容
  iframeDoc.write(printHTML);

  console.log('打印内容', printHTML);

  iframeDoc.close();
  iframe.contentWindow.focus();

  iframe.contentWindow.addEventListener('load', function() {
    iframe.contentWindow.print();
    document.body.removeChild(iframe);
  });
};