/* =========================================
   VOCABULARY HUB
   GOOGLE APPS SCRIPT
========================================= */


/* =========================================
   WEB APP
========================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Vocabulary Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}


/* =========================================
   INCLUDE HTML FILES
========================================= */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =========================================
   GET VOCABULARY
========================================= */

function getVocabulary() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Vocabulary');

  if (!sheet) {
    throw new Error('Vocabulary sheet not found.');
  }


  const data =
    sheet.getDataRange().getDisplayValues();


  if (data.length <= 1) {
    return [];
  }


  /* =========================================
     READ HEADERS
  ========================================= */

  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim()
        .replace(/^\uFEFF/, '');

    });


  /* =========================================
     FIND COLUMNS
  ========================================= */

  const columns = {};

  headers.forEach(function(header, index) {

    const key =
      header
        .toLowerCase()
        .replace(/[^a-z]/g, '');

    columns[key] = index;

  });


  /* =========================================
     CREATE VOCABULARY
  ========================================= */

  return data
    .slice(1)

    .filter(function(row) {

      return row.some(function(cell) {

        return String(cell).trim() !== '';

      });

    })

    .map(function(row) {

      return {

        ID:
          getCellValue_(row, columns.id),

        Category:
          getCellValue_(row, columns.category),

        Level:
          getCellValue_(row, columns.level),

        Word:
          getCellValue_(row, columns.word),

        Meaning:
          getCellValue_(row, columns.meaning),

        Phonics:
          getCellValue_(row, columns.phonics),

        IPA:
          getCellValue_(row, columns.ipa),

        Syllable:
          getCellValue_(row, columns.syllable),

        Image:
          getCellValue_(row, columns.image),

        Favorite:
          getCellValue_(row, columns.favorite)

      };

    });

}


/* =========================================
   GET CELL VALUE
========================================= */

function getCellValue_(row, columnIndex) {

  if (
    columnIndex === undefined ||
    columnIndex === -1
  ) {
    return '';
  }

  return String(
    row[columnIndex] || ''
  ).trim();

}

/* =========================================
   GET SETTINGS
========================================= */

function getSettings() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Spreadsheet not found.');
  }


  // Find Settings sheet safely
  const sheets = ss.getSheets();

  let sheet = null;

  for (let i = 0; i < sheets.length; i++) {

    const name =
      sheets[i].getName().trim().toLowerCase();

    if (name === 'settings') {

      sheet = sheets[i];

      break;

    }

  }


  if (!sheet) {

    throw new Error(
      'Settings sheet not found. Available sheets: ' +
      sheets
        .map(function(s) {
          return s.getName();
        })
        .join(', ')
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return {

      categories: [],

      levels: []

    };

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getDisplayValues();


  // Categories from Column A
  const categories = [
    ...new Set(

      data
        .map(function(row) {

          return String(row[0]).trim();

        })
        .filter(function(value) {

          return value !== '';

        })

    )
  ];


  // Levels from Column B
  const levels = [
    ...new Set(

      data
        .map(function(row) {

          return String(row[1]).trim();

        })
        .filter(function(value) {

          return value !== '';

        })

    )
  ];


  return {

    categories: categories,

    levels: levels

  };

}

/* =========================================
   UPDATE FAVORITE
========================================= */

function updateFavorite(id, value) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Vocabulary');

  if (!sheet) {
    throw new Error('Vocabulary sheet not found.');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return false;
  }

  const headers = data[0];

  const idColumn =
    headers.indexOf('ID');

  const favoriteColumn =
    headers.indexOf('Favorite');

  if (
    idColumn === -1 ||
    favoriteColumn === -1
  ) {
    throw new Error(
      'ID or Favorite column not found.'
    );
  }

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][idColumn]) ===
      String(id)
    ) {

      sheet
        .getRange(
          i + 1,
          favoriteColumn + 1
        )
        .setValue(value ? 'Yes' : '');

      return true;

    }

  }

  return false;

}

/* =========================================
   DELETE VOCABULARY WORD
========================================= */

function deleteVocabulary(id) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Vocabulary');

  if (!sheet) {
    throw new Error('Vocabulary sheet not found.');
  }


  const data =
    sheet.getDataRange().getValues();


  if (data.length < 2) {
    throw new Error('No vocabulary words found.');
  }


  const headers =
    data[0].map(function(header) {
      return String(header).trim();
    });


  const idColumn =
    headers.indexOf('ID');


  if (idColumn === -1) {
    throw new Error('ID column not found.');
  }


  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][idColumn]).trim() ===
      String(id).trim()
    ) {

      sheet.deleteRow(i + 1);

      return true;

    }

  }


  throw new Error(
    'Word with ID ' + id + ' was not found.'
  );

}

/* =========================================
   UPDATE VOCABULARY WORD
========================================= */

function updateVocabulary(word) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Vocabulary');

  if (!sheet) {
    throw new Error('Vocabulary sheet not found.');
  }


  const data =
    sheet.getDataRange().getValues();


  if (data.length < 2) {
    throw new Error('No vocabulary words found.');
  }


  const headers =
    data[0].map(function(header) {
      return String(header).trim();
    });


  const idColumn =
    headers.indexOf('ID');


  if (idColumn === -1) {
    throw new Error('ID column not found.');
  }


  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][idColumn]).trim() ===
      String(word.ID).trim()
    ) {

      const fields = [
        'Word',
        'Category',
        'Level',
        'Meaning',
        'Phonics',
        'IPA',
        'Syllable',
        'Image',
        'Favorite'
      ];


      fields.forEach(function(field) {

        const column =
          headers.indexOf(field);


        if (column !== -1) {

          sheet
            .getRange(i + 1, column + 1)
            .setValue(word[field] || '');

        }

      });


      return true;

    }

  }


  throw new Error(
    'Word with ID ' +
    word.ID +
    ' was not found.'
  );

}

/* =========================================
   ADD VOCABULARY WORD
========================================= */

function addVocabulary(word) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Vocabulary');


  if (!sheet) {

    throw new Error(
      'Vocabulary sheet not found.'
    );

  }


  const data =
    sheet.getDataRange().getValues();


  const headers =
    data[0].map(function(header) {

      return String(header).trim();

    });


  const idColumn =
    headers.indexOf('ID');


  if (idColumn === -1) {

    throw new Error(
      'ID column not found.'
    );

  }


  /* =====================================
     GENERATE NEW ID
  ====================================== */

  /* =====================================
   FIND LOWEST AVAILABLE ID
===================================== */

const usedIDs = new Set();

for (let i = 1; i < data.length; i++) {

  const id =
    Number(data[i][idColumn]);

  if (
    Number.isInteger(id) &&
    id > 0
  ) {

    usedIDs.add(id);

  }

}


let newID = 1;

while (usedIDs.has(newID)) {

  newID++;

}


  /* =====================================
     PREPARE NEW ROW
  ====================================== */

  const newRow =
    new Array(headers.length).fill('');


  const values = {

    ID: newID,

    Word: word.Word || '',

    Category: word.Category || '',

    Level: word.Level || '',

    Meaning: word.Meaning || '',

    Phonics: word.Phonics || '',

    IPA: word.IPA || '',

    Syllable: word.Syllable || '',

    Image: word.Image || '',

    Favorite: ''

  };


  headers.forEach(function(header, index) {

    if (
      Object.prototype.hasOwnProperty
        .call(values, header)
    ) {

      newRow[index] =
        values[header];

    }

  });


  /* =====================================
     ADD ROW
  ====================================== */

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      newRow.length
    )
    .setValues([newRow]);


  /* =====================================
     RETURN WORD TO APP
  ====================================== */

  return {

    ID: String(newID),

    Word: word.Word || '',

    Category: word.Category || '',

    Level: word.Level || '',

    Meaning: word.Meaning || '',

    Phonics: word.Phonics || '',

    IPA: word.IPA || '',

    Syllable: word.Syllable || '',

    Image: word.Image || '',

    Favorite: ''

  };

}

/* =========================================
   UPLOAD VOCABULARY IMAGE TO GITHUB
========================================= */

function uploadVocabularyImage(fileData, word, category) {

  if (!fileData || !fileData.base64) {
    throw new Error('Please select an image.');
  }

  if (!word || !word.trim()) {
    throw new Error('Please enter the word first.');
  }


  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty('GITHUB_TOKEN');

  if (!token) {
    throw new Error('GITHUB_TOKEN not found in Script Properties.');
  }


  const owner = 'rajinikanthhc';
  const repo = 'images';
  const branch = 'main';


  // Clean word for filename
  const cleanWord =
  word
    .trim()
    .replace(/[\\/:*?"<>|]/g, '')
    .replace(/\s+/g, ' ');

const extension =
  getImageExtension_(fileData.mimeType);

const fileName =
  cleanWord + extension;

const path =
  'vocabulary/' + fileName;

  const apiUrl =
    'https://api.github.com/repos/' +
    owner +
    '/' +
    repo +
    '/contents/' +
    path;


  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json'
  };


  /*
   * Check whether file already exists.
   */
  let sha = null;

  const existing =
    UrlFetchApp.fetch(
      apiUrl + '?ref=' + encodeURIComponent(branch),
      {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true
      }
    );


  if (existing.getResponseCode() === 200) {

    const existingData =
      JSON.parse(existing.getContentText());

    sha = existingData.sha;

  }


  const payload = {

    message:
      (sha ? 'Update' : 'Add') +
      ' vocabulary image: ' +
      fileName,

    content: fileData.base64,

    branch: branch

  };


  if (sha) {
    payload.sha = sha;
  }


  const response =
    UrlFetchApp.fetch(
      apiUrl,
      {
        method: 'put',
        headers: headers,
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );


  const code =
    response.getResponseCode();


  if (code !== 200 && code !== 201) {

    throw new Error(
      'GitHub upload failed: ' +
      response.getContentText()
    );

  }


  return {
  path: path,

  fileName: fileName,

  url:
    'https://raw.githubusercontent.com/' +
    owner +
    '/' +
    repo +
    '/' +
    branch +
    '/' +
    path
};

}


/* =========================================
   IMAGE EXTENSION
========================================= */

function getImageExtension_(mimeType) {

  const map = {

    'image/jpeg': '.jpg',

    'image/jpg': '.jpg',

    'image/png': '.png',

    'image/webp': '.webp',

    'image/gif': '.gif'

  };


  return map[mimeType] || '.png';

}