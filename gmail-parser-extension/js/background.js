// ============================================
// BACKGROUND.JS - Service Worker
// ============================================

importScripts('xlsx.full.min.js');

// ============================================
// 1. OAuth и работа с Gmail (с настройками пользователя)
// ============================================

function getAuthToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(token);
    });
  });
}

async function searchEmails(token, fromEmail) {
  const query = `from:${fromEmail}`;
  const url = `https://www.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(query)}&maxResults=20`;

  console.log(`   🔍 Поиск: ${query}`);

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`   ❌ Ошибка API: ${response.status} - ${errorText}`);
    throw new Error(`Gmail API error: ${response.status}`);
  }

  const data = await response.json();
  console.log(`   📊 Найдено писем: ${data.messages?.length || 0}`);
  return data.messages || [];
}

async function getMessage(token, messageId) {
  const url = `https://www.googleapis.com/gmail/v1/users/me/messages/${messageId}?format=full`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch message: ${response.status}`);
  }

  return await response.json();
}

async function getAttachment(token, messageId, attachmentId) {
  const url = `https://www.googleapis.com/gmail/v1/users/me/messages/${messageId}/attachments/${attachmentId}`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch attachment: ${response.status}`);
  }

  const data = await response.json();
  const base64 = data.data.replace(/-/g, '+').replace(/_/g, '/');
  return base64;
}

// ============================================
// 2. СКАЧИВАНИЕ ФАЙЛА
// ============================================

function downloadBase64File(
  base64Data,
  filename,
  mimeType = 'application/octet-stream',
) {
  return new Promise((resolve, reject) => {
    try {
      const dataUri = `data:${mimeType};base64,${base64Data}`;
      chrome.downloads.download(
        {
          url: dataUri,
          filename: filename,
          conflictAction: 'uniquify',
          saveAs: false,
        },
        (downloadId) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
          } else {
            resolve(downloadId);
          }
        },
      );
    } catch (error) {
      reject(error);
    }
  });
}

// ============================================
// 3. ПАРСИНГ EXCEL С ГИБКИМИ КОЛОНКАМИ
// ============================================

function parseExcelFromBase64(base64Data) {
  try {
    const byteCharacters = atob(base64Data);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);

    const workbook = XLSX.read(byteArray, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(firstSheet);

    console.log(`   📊 Парсинг: ${data.length} строк`);

    if (data.length > 0) {
      console.log(`   📋 Колонки: ${Object.keys(data[0]).join(', ')}`);
    }

    return data;
  } catch (error) {
    console.error('   ❌ Ошибка парсинга:', error.message);
    return null;
  }
}

// ============================================
// 4. ОБЪЕДИНЕНИЕ ПРАЙСОВ (С ГИБКИМИ КОЛОНКАМИ)
// ============================================

function mergePriceLists(allData, columnMapping) {
  console.log(`\n🔄 Объединение ${allData.length} файлов...`);
  console.log(`📋 Настройки колонок:`, columnMapping);

  const merged = {};

  for (const data of allData) {
    if (!data || !data.rows) {
      console.log(`   ⚠️ Пропущен файл без данных`);
      continue;
    }

    console.log(`   📦 ${data.supplier}: ${data.rows.length} строк`);

    for (const row of data.rows) {
      // ============================================
      // ИСПОЛЬЗУЕМ НАСТРОЙКИ КОЛОНОК ОТ ПОЛЬЗОВАТЕЛЯ
      // ============================================
      const brand =
        row[columnMapping.brand] ||
        row['Бренд'] ||
        row['Brand'] ||
        row['Марка'] ||
        '';
      const article =
        row[columnMapping.article] ||
        row['Артикул'] ||
        row['Article'] ||
        row['Номер'] ||
        row['Кузов'] ||
        '';
      const name =
        row[columnMapping.name] || row['Наименование'] || row['Name'] || '';
      const price = parseFloat(
        row[columnMapping.price] ||
          row['Цена'] ||
          row['Price'] ||
          row['Стоимость'] ||
          0,
      );
      const quantity = parseFloat(
        row[columnMapping.quantity] ||
          row['Количество'] ||
          row['Quantity'] ||
          row['Остаток'] ||
          0,
      );

      if (!brand && !article) {
        continue;
      }

      const key = `${brand.trim().toUpperCase()}|${article.trim().toUpperCase()}`;

      if (!merged[key]) {
        merged[key] = {
          brand: brand.trim(),
          article: article.trim(),
          name: name || `${brand} ${article}`,
          suppliers: [],
          maxPrice: 0,
          totalQuantity: 0,
        };
      }

      if (!merged[key].suppliers.includes(data.supplier)) {
        merged[key].suppliers.push(data.supplier);
      }

      if (price > merged[key].maxPrice) {
        merged[key].maxPrice = price;
      }

      merged[key].totalQuantity += quantity || 0;
    }
  }

  const result = Object.values(merged).map((item) => ({
    Бренд: item.brand,
    Артикул: item.article,
    Наименование: item.name,
    'Макс. цена': item.maxPrice,
    'Общее кол-во': item.totalQuantity,
    Поставщики: item.suppliers.join(', '),
  }));

  result.sort((a, b) => {
    if (a['Бренд'] !== b['Бренд']) return a['Бренд'].localeCompare(b['Бренд']);
    return a['Артикул'].localeCompare(b['Артикул']);
  });

  console.log(`   📊 Итог: ${result.length} уникальных позиций`);
  return result;
}

// ============================================
// 5. СОЗДАНИЕ ИТОГОВОГО EXCEL
// ============================================

function createMergedExcel(mergedData) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(mergedData);

  ws['!cols'] = [
    { wch: 20 }, // Бренд
    { wch: 20 }, // Артикул
    { wch: 40 }, // Наименование
    { wch: 15 }, // Макс. цена
    { wch: 15 }, // Общее кол-во
    { wch: 30 }, // Поставщики
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Общий прайс');
  return wb;
}

function excelToBase64(wb) {
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return btoa(
    new Uint8Array(wbOut).reduce(
      (data, byte) => data + String.fromCharCode(byte),
      '',
    ),
  );
}

// ============================================
// 6. ПОИСК ВЛОЖЕНИЙ
// ============================================

function findAttachments(parts) {
  const attachments = [];
  if (!parts) return attachments;

  for (const part of parts) {
    if (part.filename && part.filename.length > 0) {
      const ext = part.filename.split('.').pop().toLowerCase();
      if (['xlsx', 'xls'].includes(ext)) {
        attachments.push({
          filename: part.filename,
          attachmentId: part.body?.attachmentId,
          mimeType: part.mimeType || 'application/octet-stream',
          ext: ext,
        });
      }
    }
    if (part.parts) {
      attachments.push(...findAttachments(part.parts));
    }
  }
  return attachments;
}

// ============================================
// 7. ИСТОРИЯ
// ============================================

async function getDownloadHistory() {
  const data = await chrome.storage.local.get(['downloadHistory']);
  return data.downloadHistory || {};
}

async function isAlreadyDownloaded(messageId, filename) {
  const history = await getDownloadHistory();
  return !!history[`${messageId}_${filename}`];
}

async function addToDownloadHistory(messageId, filename, supplierName) {
  const history = await getDownloadHistory();
  const key = `${messageId}_${filename}`;
  history[key] = {
    messageId: messageId,
    filename: filename,
    supplier: supplierName,
    downloadedAt: new Date().toISOString(),
  };
  await chrome.storage.local.set({ downloadHistory: history });
}

// ============================================
// 8. ГЛАВНАЯ ФУНКЦИЯ
// ============================================

async function parseEmails(config, emailSettings, columnMapping) {
  console.log('🚀 Запуск парсинга...');
  console.log('📋 Поставщиков:', config.suppliers.length);
  console.log('📋 Колонки:', columnMapping);

  if (!config || !config.suppliers || config.suppliers.length === 0) {
    throw new Error('Список поставщиков пуст');
  }

  const results = {
    processed: 0,
    filesProcessed: 0,
    skipped: 0,
    errors: [],
    mergedCount: 0,
  };

  try {
    // Используем OAuth (Gmail API) — пароль пользователя не передается в код!
    const token = await getAuthToken();
    console.log('✅ Получен OAuth-токен');

    const allSupplierData = [];

    for (const supplier of config.suppliers) {
      if (!supplier || !supplier.email) {
        console.warn('⚠️ Пропущен поставщик без email');
        continue;
      }

      console.log(`\n📦 Поставщик: ${supplier.name} (${supplier.email})`);

      try {
        const messages = await searchEmails(token, supplier.email);

        if (messages.length === 0) {
          console.log(`   ℹ️ Нет писем`);
          continue;
        }

        console.log(`   📨 Найдено ${messages.length} писем`);

        for (const msg of messages) {
          console.log(`   📨 Письмо ID: ${msg.id}`);

          const message = await getMessage(token, msg.id);
          const parts = message.payload?.parts || [];
          const attachments = findAttachments(parts);

          if (attachments.length === 0) {
            console.log(`   ⚠️ Нет Excel-вложений`);
            continue;
          }

          for (const attach of attachments) {
            try {
              if (!attach.attachmentId) {
                console.log(`   ⚠️ Нет attachmentId`);
                continue;
              }

              if (await isAlreadyDownloaded(msg.id, attach.filename)) {
                console.log(
                  `   ⏭️ Пропускаем ${attach.filename} (уже обработан)`,
                );
                results.skipped++;
                continue;
              }

              console.log(`   📥 Обработка: ${attach.filename}...`);

              const base64Data = await getAttachment(
                token,
                msg.id,
                attach.attachmentId,
              );

              // Сохраняем оригинал
              const originalFilename = `${supplier.name}_${attach.filename}`;
              await downloadBase64File(
                base64Data,
                originalFilename,
                attach.mimeType,
              );
              console.log(`   💾 Сохранен: ${originalFilename}`);

              // Парсим данные
              const excelData = parseExcelFromBase64(base64Data);

              if (excelData && excelData.length > 0) {
                allSupplierData.push({
                  supplier: supplier.name,
                  filename: attach.filename,
                  rows: excelData,
                });
                results.filesProcessed++;
                console.log(`   ✅ Добавлено ${excelData.length} строк`);
              }

              await addToDownloadHistory(
                msg.id,
                attach.filename,
                supplier.name,
              );
            } catch (err) {
              console.error(`   ❌ Ошибка: ${err.message}`);
              results.errors.push({
                supplier: supplier.name,
                error: err.message,
              });
            }
          }
          results.processed++;
        }
      } catch (err) {
        console.error(`   ❌ Ошибка: ${err.message}`);
        results.errors.push({ supplier: supplier.email, error: err.message });
      }
    }

    // ============================================
    // ОБЪЕДИНЕНИЕ С НАСТРОЙКАМИ КОЛОНОК
    // ============================================
    console.log(`\n📊 Собрано данных: ${allSupplierData.length} файлов`);

    if (allSupplierData.length > 0) {
      const mergedData = mergePriceLists(allSupplierData, columnMapping);

      if (mergedData && mergedData.length > 0) {
        const wb = createMergedExcel(mergedData);
        const base64 = excelToBase64(wb);

        const filename = `Общий_прайс_${new Date().toISOString().slice(0, 10)}.xlsx`;
        await downloadBase64File(base64, filename);

        results.mergedCount = mergedData.length;
        console.log(
          `💾 Итоговый прайс: ${filename} (${mergedData.length} позиций)`,
        );
      }
    }

    console.log('\n📊 ИТОГИ:');
    console.log(`   📁 Обработано файлов: ${results.filesProcessed}`);
    console.log(`   📋 Итоговый прайс: ${results.mergedCount} позиций`);

    return results;
  } catch (error) {
    console.error('❌ Критическая ошибка:', error.message);
    throw error;
  }
}

// ============================================
// 9. ОБРАБОТКА СООБЩЕНИЙ
// ============================================

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('📩 Получено сообщение:', request.action);

  if (request.action === 'startParsing') {
    if (!request.config || !request.emailSettings || !request.columnMapping) {
      sendResponse({
        success: false,
        error: 'Не все настройки переданы',
      });
      return true;
    }

    parseEmails(request.config, request.emailSettings, request.columnMapping)
      .then((result) => {
        console.log('✅ Парсинг завершен');
        sendResponse({ success: true, data: result });
      })
      .catch((error) => {
        console.error('❌ Ошибка:', error.message);
        sendResponse({ success: false, error: error.message });
      });
    return true;
  }

  if (request.action === 'getStatus') {
    sendResponse({ status: 'ready' });
    return true;
  }

  if (request.action === 'clearHistory') {
    chrome.storage.local.set({ downloadHistory: {} }, () => {
      console.log('🗑️ История очищена');
      sendResponse({ success: true });
    });
    return true;
  }

  if (request.action === 'getHistoryStats') {
    getDownloadHistory().then((history) => {
      sendResponse({ count: Object.keys(history).length });
    });
    return true;
  }
});

// ============================================
// 10. ПРИ УСТАНОВКЕ
// ============================================

chrome.runtime.onInstalled.addListener(() => {
  console.log('✅ Gmail Parser установлен');
  chrome.storage.local.set({
    config: { suppliers: [] },
    downloadHistory: {},
  });
});
