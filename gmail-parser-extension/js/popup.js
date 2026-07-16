// ============================================
// POPUP.JS
// ============================================

// Загрузка сохраненных настроек
document.addEventListener('DOMContentLoaded', async () => {
  try {
    const data = await chrome.storage.local.get([
      'config',
      'emailSettings',
      'columnMapping',
    ]);

    // Загружаем настройки почты
    if (data.emailSettings) {
      document.getElementById('email').value = data.emailSettings.email || '';
      document.getElementById('password').value =
        data.emailSettings.password || '';
      document.getElementById('imapServer').value =
        data.emailSettings.imapServer || 'imap.gmail.com';
      document.getElementById('port').value = data.emailSettings.port || 993;
    }

    // Загружаем настройки колонок
    if (data.columnMapping) {
      document.getElementById('colBrand').value =
        data.columnMapping.brand || 'Бренд';
      document.getElementById('colArticle').value =
        data.columnMapping.article || 'Артикул';
      document.getElementById('colName').value =
        data.columnMapping.name || 'Наименование';
      document.getElementById('colPrice').value =
        data.columnMapping.price || 'Цена';
      document.getElementById('colQuantity').value =
        data.columnMapping.quantity || 'Количество';
    }

    // Загружаем поставщиков
    const list = document.getElementById('supplierList');
    list.innerHTML = '';

    if (
      data.config &&
      data.config.suppliers &&
      data.config.suppliers.length > 0
    ) {
      for (const supplier of data.config.suppliers) {
        addSupplier(supplier.name, supplier.email);
      }
    } else {
      addSupplier();
    }

    addLog('📂 Настройки загружены', 'system');
  } catch (error) {
    addLog(`❌ Ошибка: ${error.message}`, 'error');
  }

  // ============================================
  // ОБРАБОТЧИКИ СОБЫТИЙ
  // ============================================

  // Добавление поставщика
  document.getElementById('addSupplier').addEventListener('click', function () {
    addSupplier();
    addLog('➕ Добавлен новый поставщик', 'system');
  });

  // Сохранение настроек
  document
    .getElementById('saveBtn')
    .addEventListener('click', async function () {
      try {
        const config = getConfigFromUI();
        const emailSettings = getEmailSettings();
        const columnMapping = getColumnMapping();

        if (!emailSettings.email || !emailSettings.password) {
          addLog('❌ Введите email и пароль приложения!', 'error');
          return;
        }

        if (!config.suppliers || config.suppliers.length === 0) {
          addLog('❌ Добавьте хотя бы одного поставщика!', 'error');
          return;
        }

        await chrome.storage.local.set({
          config,
          emailSettings,
          columnMapping,
        });

        addLog('✅ Настройки сохранены!', 'success');
      } catch (error) {
        addLog(`❌ Ошибка сохранения: ${error.message}`, 'error');
      }
    });

  // Запуск парсинга
  document
    .getElementById('startBtn')
    .addEventListener('click', async function () {
      try {
        const config = getConfigFromUI();
        const emailSettings = getEmailSettings();
        const columnMapping = getColumnMapping();

        if (!emailSettings.email || !emailSettings.password) {
          addLog('❌ Введите email и пароль приложения!', 'error');
          return;
        }

        if (!config.suppliers || config.suppliers.length === 0) {
          addLog('❌ Добавьте хотя бы одного поставщика!', 'error');
          return;
        }

        if (
          !columnMapping.brand ||
          !columnMapping.article ||
          !columnMapping.price
        ) {
          addLog('❌ Укажите названия колонок: Бренд, Артикул, Цена!', 'error');
          return;
        }

        addLog(
          `🚀 Запуск парсинга для ${config.suppliers.length} поставщиков...`,
          'info',
        );
        updateStatus('running', 'Идет обработка...');

        await chrome.storage.local.set({
          config,
          emailSettings,
          columnMapping,
        });

        const response = await chrome.runtime.sendMessage({
          action: 'startParsing',
          config: config,
          emailSettings: emailSettings,
          columnMapping: columnMapping,
        });

        if (response && response.success) {
          const data = response.data;
          addLog(
            `✅ Обработано файлов: ${data.filesProcessed || 0}`,
            'success',
          );
          if (data.mergedCount > 0) {
            addLog(`📋 Итоговый прайс: ${data.mergedCount} позиций`, 'success');
          }
          if (data.errors && data.errors.length > 0) {
            addLog(`⚠️ Ошибок: ${data.errors.length}`, 'warning');
          }
          updateStatus('ready', 'Готов');
        } else {
          addLog(
            `❌ Ошибка: ${response?.error || 'Неизвестная ошибка'}`,
            'error',
          );
          updateStatus('error', 'Ошибка');
        }
      } catch (error) {
        console.error('❌ Ошибка:', error);
        addLog(`❌ Критическая ошибка: ${error.message}`, 'error');
        updateStatus('error', 'Ошибка');
      }
    });

  // ============================================
  // ПОЛНАЯ ОЧИСТКА ВСЕХ ДАННЫХ
  // ============================================
document.getElementById('clearBtn').addEventListener('click', function () {
  if (
    confirm(
      '⚠️ Вы уверены, что хотите удалить ВСЕ настройки и данные?\n\nБудут удалены:\n• Email и пароль\n• Настройки IMAP и колонок\n• Список поставщиков\n• Кэш авторизации Google',
    )
  ) {
    // 1. Очищаем поля ввода (УБИРАЕМ ВСЁ!)
    document.getElementById('email').value = '';
    document.getElementById('password').value = '';
    document.getElementById('imapServer').value = 'imap.gmail.com';
    document.getElementById('port').value = '993';

    document.getElementById('colBrand').value = 'Бренд';
    document.getElementById('colArticle').value = 'Артикул';
    document.getElementById('colName').value = 'Наименование';
    document.getElementById('colPrice').value = 'Цена';
    document.getElementById('colQuantity').value = 'Количество';

    // Поставщики
    const list = document.getElementById('supplierList');
    list.innerHTML = '';
    addSupplier();

    // 2. Очищаем хранилище ПОЛНОСТЬЮ
    chrome.storage.local.clear(function () {
      console.log('✅ chrome.storage.local очищен');
    });

    // 3. Очищаем кэш авторизации Google
    chrome.identity.clearAllCachedAuthTokens(function () {
      console.log('✅ Кэш Google очищен');
    });

    // 4. Устанавливаем флаг, что был полный сброс
    chrome.storage.local.set({ _resetDone: true });

    addLog('🗑️ ВСЕ НАСТРОЙКИ И ДАННЫЕ УДАЛЕНЫ', 'system');
    addLog('✅ Очищено: хранилище, кэш Google, поля ввода', 'success');
    addLog('📝 Введите новые настройки и нажмите "Сохранить"', 'info');

    updateStatus('ready', 'Полный сброс выполнен');
  }
});
});

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

function addSupplier(name = '', email = '') {
  const list = document.getElementById('supplierList');
  const item = document.createElement('div');
  item.className = 'supplier-item';
  item.innerHTML = `
        <input class="supplier-name" placeholder="Имя" value="${name}">
        <input class="supplier-email" placeholder="email@example.com" value="${email}">
        <button class="btn-remove">✕</button>
    `;

  const removeBtn = item.querySelector('.btn-remove');
  removeBtn.addEventListener('click', function () {
    const items = document.querySelectorAll('.supplier-item');
    if (items.length <= 1) {
      addLog('❌ Должен быть хотя бы один поставщик', 'error');
      return;
    }
    item.remove();
    addLog('🗑️ Поставщик удален', 'system');
  });

  list.appendChild(item);
}

function getConfigFromUI() {
  const items = document.querySelectorAll('.supplier-item');
  const suppliers = [];
  for (const item of items) {
    const name = item.querySelector('.supplier-name').value.trim();
    const email = item.querySelector('.supplier-email').value.trim();
    if (name && email) {
      suppliers.push({ name, email });
    }
  }
  return { suppliers };
}

function getEmailSettings() {
  return {
    email: document.getElementById('email').value.trim(),
    password: document.getElementById('password').value.trim(),
    imapServer:
      document.getElementById('imapServer').value.trim() || 'imap.gmail.com',
    port: parseInt(document.getElementById('port').value) || 993,
  };
}

function getColumnMapping() {
  return {
    brand: document.getElementById('colBrand').value.trim() || 'Бренд',
    article: document.getElementById('colArticle').value.trim() || 'Артикул',
    name: document.getElementById('colName').value.trim() || 'Наименование',
    price: document.getElementById('colPrice').value.trim() || 'Цена',
    quantity:
      document.getElementById('colQuantity').value.trim() || 'Количество',
  };
}

function addLog(message, type = 'info') {
  const container = document.getElementById('logContainer');
  const entry = document.createElement('div');
  entry.className = `log-entry ${type}`;
  entry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
  container.appendChild(entry);
  container.scrollTop = container.scrollHeight;
}

function updateStatus(status, message) {
  const dot = document.getElementById('statusBar').querySelector('.status-dot');
  const text = document.getElementById('statusText');
  dot.className = `status-dot ${status}`;
  text.textContent = message;
}
