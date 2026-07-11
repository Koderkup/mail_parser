const fs = require('fs');
const path = require('path');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

// ============================================
// ЧИТАЕМ КОНФИГУРАЦИОННЫЙ ФАЙЛ
// ============================================
let config;
try {
  const configPath = path.join(__dirname, 'config.json');
  const configContent = fs.readFileSync(configPath, 'utf8');
  config = JSON.parse(configContent);
  console.log(`📧 Почтовый ящик: ${config.emailSettings.email}`);
  console.log(`📋 Поставщиков: ${config.suppliers.length}`);
} catch (error) {
  console.error('❌ Ошибка загрузки config.json:', error.message);
  console.log(
    '⚠️  Убедитесь, что файл config.json существует и имеет правильный формат.',
  );

  process.exit(1);
}

// ============================================
//  ФУНКЦИЯ ДЛЯ ОБРАБОТКИ EXCEL-ФАЙЛОВ
// ============================================
function processExcelFile(filePath) {
  try {
    console.log(`   📊 Чтение файла: ${path.basename(filePath)}`);
    const workbook = XLSX.readFile(filePath);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(firstSheet);
    console.log(`   📋 Найдено ${data.length} строк`);
    return data;
  } catch (error) {
    console.error(`   ❌ Ошибка чтения Excel: ${error.message}`);
    return null;
  }
}

// ============================================
//  ФУНКЦИЯ ДЛЯ РАСПАКОВКИ АРХИВОВ
// ============================================
function extractArchive(filePath) {
  try {
    const ext = path.extname(filePath).toLowerCase();
    const dir = path.dirname(filePath);
    const baseName = path.basename(filePath, ext);

    if (ext === '.zip') {
      console.log(`   📦 Распаковка ZIP: ${path.basename(filePath)}`);
      const zip = new AdmZip(filePath);
      const extractedPath = path.join(dir, `${baseName}_extracted`);
      if (!fs.existsSync(extractedPath)) {
        fs.mkdirSync(extractedPath);
      }
      zip.extractAllTo(extractedPath, true);

      // Ищем XLSX файлы в распакованной папке
      const files = fs.readdirSync(extractedPath);
      for (let file of files) {
        const fileExt = path.extname(file).toLowerCase();
        if (fileExt === '.xlsx' || fileExt === '.xls') {
          const fullPath = path.join(extractedPath, file);
          console.log(`   📄 Найден Excel в архиве: ${file}`);
          return processExcelFile(fullPath);
        }
      }
      return null;
    } else if (ext === '.cab') {
      console.log(`   📦 Распаковка CAB: ${path.basename(filePath)}`);
      // Для CAB используем системную утилиту expand.exe
      const { execSync } = require('child_process');
      const targetDir = path.join(dir, `${baseName}_extracted`);
      if (!fs.existsSync(targetDir)) {
        fs.mkdirSync(targetDir);
      }
      try {
        execSync(`expand "${filePath}" -F:* "${targetDir}"`, { stdio: 'pipe' });
        console.log(`   ✅ CAB распакован в: ${targetDir}`);

        // Ищем XLSX файлы в распакованной папке
        const files = fs.readdirSync(targetDir);
        for (let file of files) {
          const fileExt = path.extname(file).toLowerCase();
          if (fileExt === '.xlsx' || fileExt === '.xls') {
            const fullPath = path.join(targetDir, file);
            console.log(`   📄 Найден Excel в CAB: ${file}`);
            return processExcelFile(fullPath);
          }
        }
      } catch (err) {
        console.error(`   ❌ Ошибка распаковки CAB: ${err.message}`);
      }
      return null;
    } else {
      console.log(`   ⚠️  Неизвестный формат архива: ${ext}`);
      return null;
    }
  } catch (error) {
    console.error(`   ❌ Ошибка обработки архива: ${error.message}`);
    return null;
  }
}

// ============================================
// 5. ФУНКЦИЯ ДЛЯ СКАЧИВАНИЯ И ОБРАБОТКИ ВЛОЖЕНИЙ
// ============================================
function downloadAttachmentsFromImap(supplier) {
  return new Promise((resolve, reject) => {
    const imapConfig = {
      user: config.emailSettings.email,
      password: config.emailSettings.password,
      host: config.emailSettings.imapServer,
      port: config.emailSettings.port,
      tls: true,
      tlsOptions: { rejectUnauthorized: false },
      authTimeout: 30000,
    };

    const imap = new Imap(imapConfig);
    let downloadedCount = 0;
    let isResolved = false;

    const safeResolve = (result) => {
      if (!isResolved) {
        isResolved = true;
        resolve(result);
      }
    };

    const safeReject = (error) => {
      if (!isResolved) {
        isResolved = true;
        reject(error);
      }
    };

    imap.once('ready', function () {
      console.log(`   ✅ Подключение к ${supplier.email}...`);

      imap.openBox('INBOX', false, function (err) {
        if (err) {
          imap.end();
          safeReject(err);
          return;
        }

        const searchCriteria = ['UNSEEN', ['FROM', supplier.email]];
        imap.search(searchCriteria, function (err, results) {
          if (err) {
            imap.end();
            safeReject(err);
            return;
          }

          if (results.length === 0) {
            console.log(`   ℹ️  Новых писем от ${supplier.name} не найдено.`);
            imap.end();
            safeResolve(0);
            return;
          }

          console.log(
            `   📨 Найдено ${results.length} новых писем от ${supplier.name}`,
          );

          const fetch = imap.fetch(results, {
            bodies: [''],
            struct: true,
          });

          let processedCount = 0;

          fetch.on('message', function (msg) {
            msg.on('body', function (stream) {
              simpleParser(stream, (err, parsed) => {
                if (err) {
                  console.error('   ❌ Ошибка парсинга письма:', err.message);
                  return;
                }

                if (!parsed.attachments || parsed.attachments.length === 0) {
                  console.log('   ⚠️  В письме нет вложений');
                  return;
                }

                console.log(
                  `   📎 Найдено ${parsed.attachments.length} вложений`,
                );

                const downloadsPath = path.join(__dirname, 'downloads');
                if (!fs.existsSync(downloadsPath)) {
                  fs.mkdirSync(downloadsPath);
                }

                // Обрабатываем каждое вложение
                for (let attach of parsed.attachments) {
                  const filePath = path.join(
                    downloadsPath,
                    `${supplier.name}_${attach.filename}`,
                  );

                  fs.writeFileSync(filePath, attach.content);
                  console.log(`   💾 Сохранен: ${filePath}`);
                  downloadedCount++;

                  // Обрабатываем файл в зависимости от расширения
                  const ext = path.extname(attach.filename).toLowerCase();
                  if (ext === '.xlsx' || ext === '.xls') {
                    processExcelFile(filePath);
                  } else if (ext === '.zip' || ext === '.cab') {
                    extractArchive(filePath);
                  }
                }
              });
            });
          });

          fetch.once('error', function (err) {
            console.error('   ❌ Ошибка загрузки:', err.message);
            imap.end();
            safeReject(err);
          });

          fetch.once('end', function () {
            console.log(`   ✅ Завершена обработка писем от ${supplier.name}`);
            // Даем время на обработку всех вложений
            setTimeout(() => {
              imap.end();
              safeResolve(downloadedCount);
            }, 2000);
          });
        });
      });
    });

    imap.once('error', function (err) {
      console.error('   ❌ Ошибка IMAP:', err.message);
      safeReject(err);
    });

    imap.once('end', function () {
      console.log(`   🔌 Соединение закрыто для ${supplier.name}`);
    });

    imap.connect();
  });
}

// ============================================
// ГЛАВНАЯ ФУНКЦИЯ ЗАПУСКА
// ============================================
async function main() {
  console.log('\n🚀 ЗАПУСК ПРОГРАММЫ ОБРАБОТКИ ПРАЙСОВ');
  console.log('='.repeat(50));

  const startTime = Date.now();

  try {
    const downloadsPath = path.join(__dirname, 'downloads');
    if (!fs.existsSync(downloadsPath)) {
      fs.mkdirSync(downloadsPath);
    }

    console.log('\n📧 Обработка писем поставщиков:');
    console.log('-'.repeat(40));

    for (let supplier of config.suppliers) {
      console.log(`\n📦 Поставщик: ${supplier.name}`);
      await downloadAttachmentsFromImap(supplier);
    }

    const elapsed = (Date.now() - startTime) / 1000;
    console.log('\n' + '='.repeat(50));
    console.log(`✅ РАБОТА ЗАВЕРШЕНА за ${elapsed.toFixed(1)} секунд`);
    console.log(`📁 Файлы сохранены в папку: ${downloadsPath}`);

    if (fs.existsSync(downloadsPath)) {
      const files = fs.readdirSync(downloadsPath);
      if (files.length > 0) {
        console.log(`📎 Скачано файлов: ${files.length}`);
        console.log(`   ${files.join('\n   ')}`);
      } else {
        console.log('📂 Папка загрузок пуста.');
      }
    }
  } catch (error) {
    console.error('\n❌ КРИТИЧЕСКАЯ ОШИБКА:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

main();
