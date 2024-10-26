/**
 *
 * Shopping Performance Оптимізатор
 *
 * Цей скрипт призначений для автоматизації аналізу продуктивності товарів у Google Ads та запису результатів у Google Sheets.
 * Він дозволяє збирати дані за будь-який вказаний період, включаючи покази, кліки, конверсії, цінність конверсій і витрати,
 * а також призначати мітки на основі цих показників.
 * Скрипт є універсальним, що дозволяє легко адаптувати його під різні бізнес-потреби.
 *
 * Google Ads Script за підримкою Igor Pecherksih
 *
 */


// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//
// Options

var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1AxKTTIUwnbsFpSmE-W6XhisP-arYtEC0AXKeoCLPgtE/edit";
var daysAgo = 90; // Період для збору статистики в днях

function main() {
  Logger.log("Початок обробки даних за останні " + daysAgo + " днів..."); // Початок логування
  var products = getFilteredProducts(daysAgo);
  Logger.log("Оброблено товарів: " + products.length); // Логування кількості товарів

  // Загальні показники
  var totalImpressions = 0;
  var totalClicks = 0;

  // Обчислюємо загальні покази та кліки
  products.forEach((product) => {
    totalImpressions += product[1]; // Покази
    totalClicks += product[2]; // Кліки
  });

  // Обчислюємо середні значення
  var avgImpressions =
    products.length > 0 ? totalImpressions / products.length : 0;
  var avgClicks = products.length > 0 ? totalClicks / products.length : 0;

  // Мітки для колонок H, I, J
  var labelsH = [];
  var labelsI = [];
  var labelsJ = [];

  // Призначаємо мітки в залежності від конверсій та показників
  products.forEach((product) => {
    // Мітки для колонки H
    labelsH.push(product[4] > 0 ? "топ товари" : "без продажів");

    // Мітки для колонки I
    labelsI.push(
      product[1] > avgImpressions ? "високі покази" : "занижені покази"
    );

    // Мітки для колонки J
    labelsJ.push(product[2] > avgClicks ? "високі кліки" : "занижені кліки");
  });

  // Записуємо дані в таблицю
  pushToSpreadsheet(products, labelsH, labelsI, labelsJ);
  Logger.log("Дані успішно записані до таблиці."); // Логування успішного запису
}

function getFilteredProducts(daysAgo) {
  var today = new Date();
  var pastDate = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate() - daysAgo
  );
  var dateFrom = Utilities.formatDate(
    pastDate,
    AdWordsApp.currentAccount().getTimeZone(),
    "yyyyMMdd"
  );
  var dateTo = Utilities.formatDate(
    today,
    AdWordsApp.currentAccount().getTimeZone(),
    "yyyyMMdd"
  );

  var query = "SELECT OfferId, Impressions, Clicks, Cost, Conversions, ConversionValue " + "FROM SHOPPING_PERFORMANCE_REPORT DURING " + dateFrom + "," + dateTo;

  var products = [];
  var report = AdWordsApp.report(query);
  var rows = report.rows();

  while (rows.hasNext()) {
    var row = rows.next();
    var offerId = row["OfferId"];
    var impressions = parseInt(row["Impressions"]) || 0;
    var clicks = parseInt(row["Clicks"]) || 0;
    var cost = parseFloat(row["Cost"].replace(",", "")) || 0;
    var conversions = parseInt(row["Conversions"]) || 0;
    var conversionValue =
      parseFloat(row["ConversionValue"].replace(",", "")) || 0;

    products.push([
      offerId,
      impressions,
      clicks,
      cost,
      conversions,
      conversionValue,
    ]);
  }
  return products;
}

function pushToSpreadsheet(data, labelsH, labelsI, labelsJ) {
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = spreadsheet.getSheetByName("Звіт товарів");
  sheet.getRange("A2:J" + sheet.getMaxRows()).clearContent(); // Очищення таблиці перед новим записом

  if (data.length > 0) {
    var range = sheet.getRange("A2:F" + (data.length + 1));
    range.setValues(data);

    sheet
      .getRange("H2:H" + (data.length + 1))
      .setValues(labelsH.map((label) => [label]));
    var costs = data.map((product) => [product[3]]);
    sheet.getRange("G2:G" + (data.length + 1)).setValues(costs);
    sheet
      .getRange("I2:I" + (data.length + 1))
      .setValues(labelsI.map((label) => [label]));
    sheet
      .getRange("J2:J" + (data.length + 1))
      .setValues(labelsJ.map((label) => [label]));
  }
}
