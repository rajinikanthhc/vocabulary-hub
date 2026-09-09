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
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

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
    throw new Error(
      'Vocabulary sheet not found.'
    );
  }

  const data =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (data.length <= 1) {
    return [];
  }


  /* =====================================
     READ HEADERS
  ====================================== */

  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim()
        .replace(/^\uFEFF/, '');

    });


  /* =====================================
     FIND COLUMNS
  ====================================== */

  const columns = {};

  headers.forEach(function(header, index) {

    const key =
      header
        .toLowerCase()
        .replace(/[^a-z]/g, '');

    columns[key] = index;

  });


  /* =====================================
     RETURN WORDS
  ====================================== */

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
          getCellValue_(
            row,
            columns.id
          ),

        Category:
          getCellValue_(
            row,
            columns.category
          ),

        Level:
          getCellValue_(
            row,
            columns.level
          ),

        Word:
          getCellValue_(
            row,
            columns.word
          ),

        Meaning:
          getCellValue_(
            row,
            columns.meaning
          ),

        Phonics:
          getCellValue_(
            row,
            columns.phonics
          ),

        IPA:
          getCellValue_(
            row,
            columns.ipa
          ),

        Syllable:
          getCellValue_(
            row,
            columns.syllable
          ),

        Image:
          getCellValue_(
            row,
            columns.image
          ),

        Favorite:
          getCellValue_(
            row,
            columns.favorite
          )

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
   UPDATE FAVORITE
========================================= */

function updateFavorite(id, value) {

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

  if (data.length < 2) {
    return false;
  }

  const headers =
    data[0];

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

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][idColumn]) ===
      String(id)
    ) {

      sheet
        .getRange(
          i + 1,
          favoriteColumn + 1
        )
        .setValue(
          value ? 'Yes' : ''
        );

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
    throw new Error(
      'Vocabulary sheet not found.'
    );
  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error(
      'No vocabulary words found.'
    );
  }

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
     FIND AND DELETE
  ====================================== */

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][idColumn]).trim() ===
      String(id).trim()
    ) {

      sheet.deleteRow(i + 1);

      return true;

    }

  }


  throw new Error(
    'Word with ID ' +
    id +
    ' was not found.'
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
    throw new Error(
      'Vocabulary sheet not found.'
    );
  }

  const data =
    sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error(
      'No vocabulary words found.'
    );
  }

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
     DUPLICATE WORD CHECK
     Excludes current ID
  ====================================== */

  const wordColumn =
    headers.indexOf('Word');

  if (wordColumn !== -1) {

    const newWord =
      String(word.Word || '')
        .trim()
        .toLowerCase();

    for (
      let i = 1;
      i < data.length;
      i++
    ) {

      const existingID =
        String(
          data[i][idColumn] || ''
        ).trim();

      const existingWord =
        String(
          data[i][wordColumn] || ''
        )
          .trim()
          .toLowerCase();

      if (
        existingID !==
          String(word.ID).trim() &&
        newWord &&
        existingWord === newWord
      ) {

        throw new Error(
          'Word already exists: "' +
          data[i][wordColumn] +
          '"'
        );

      }

    }

  }


  /* =====================================
     UPDATE ROW
  ====================================== */

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

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
            .getRange(
              i + 1,
              column + 1
            )
            .setValue(
              word[field] || ''
            );

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

  const wordColumn =
    headers.indexOf('Word');


  if (idColumn === -1) {
    throw new Error(
      'ID column not found.'
    );
  }


  if (wordColumn === -1) {
    throw new Error(
      'Word column not found.'
    );
  }


  /* =====================================
     DUPLICATE WORD CHECK
     SERVER SIDE
  ====================================== */

  const newWord =
    String(word.Word || '')
      .trim()
      .toLowerCase();

  if (!newWord) {
    throw new Error(
      'Please enter a word.'
    );
  }


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const existingWord =
      String(
        data[i][wordColumn] || ''
      )
        .trim()
        .toLowerCase();

    if (
      existingWord &&
      existingWord === newWord
    ) {

      throw new Error(
        'Word already exists: "' +
        data[i][wordColumn] +
        '"'
      );

    }

  }


  /* =====================================
     LOWEST AVAILABLE ID
  ====================================== */

  const usedIDs =
    new Set();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

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


  while (
    usedIDs.has(newID)
  ) {

    newID++;

  }


  /* =====================================
     PREPARE NEW ROW
  ====================================== */

  const newRow =
    new Array(headers.length)
      .fill('');


  const values = {

    ID: newID,

    Word:
      word.Word || '',

    Category:
      word.Category || '',

    Level:
      word.Level || '',

    Meaning:
      word.Meaning || '',

    Phonics:
      word.Phonics || '',

    IPA:
      word.IPA || '',

    Syllable:
      word.Syllable || '',

    Image:
      word.Image || '',

    Favorite:
      ''

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


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      newRow.length
    )
    .setValues([newRow]);


  return {

    ID:
      String(newID),

    Word:
      word.Word || '',

    Category:
      word.Category || '',

    Level:
      word.Level || '',

    Meaning:
      word.Meaning || '',

    Phonics:
      word.Phonics || '',

    IPA:
      word.IPA || '',

    Syllable:
      word.Syllable || '',

    Image:
      word.Image || '',

    Favorite:
      ''

  };

}


/* =========================================
   UPLOAD IMAGE TO GITHUB
========================================= */

function uploadVocabularyImage(
  fileData,
  word,
  category
) {

  if (
    !fileData ||
    !fileData.base64
  ) {

    throw new Error(
      'Image data is missing.'
    );

  }


  if (
    !word ||
    !String(word).trim()
  ) {

    throw new Error(
      'Please enter the word first.'
    );

  }


  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GITHUB_TOKEN'
      );


  if (!token) {

    throw new Error(
      'GITHUB_TOKEN not found in Script Properties.'
    );

  }


  const owner =
    'rajinikanthhc';

  const repo =
    'images';

  const branch =
    'main';


  /* =====================================
     CLEAN FILENAME
  ====================================== */

  const cleanWord =
    String(word)
      .trim()
      .replace(
        /[\\/:*?"<>|]/g,
        ''
      )
      .replace(
        /\s+/g,
        ' '
      );


  const extension =
    getImageExtension_(
      fileData.mimeType
    );


  const fileName =
    cleanWord +
    extension;


  const path =
    'vocabulary/' +
    fileName;


  const apiUrl =
    'https://api.github.com/repos/' +
    owner +
    '/' +
    repo +
    '/contents/' +
    path;


  const headers = {

    Authorization:
      'Bearer ' + token,

    Accept:
      'application/vnd.github+json'

  };


  /* =====================================
     CHECK EXISTING FILE
  ====================================== */

  let sha = null;


  const existing =
    UrlFetchApp.fetch(
      apiUrl +
      '?ref=' +
      encodeURIComponent(branch),
      {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true
      }
    );


  if (
    existing.getResponseCode() === 200
  ) {

    const existingData =
      JSON.parse(
        existing.getContentText()
      );

    sha =
      existingData.sha;

  }


  /* =====================================
     GITHUB PAYLOAD
  ====================================== */

  const payload = {

    message:
      (sha ? 'Update' : 'Add') +
      ' vocabulary image: ' +
      fileName,

    content:
      fileData.base64,

    branch:
      branch

  };


  if (sha) {

    payload.sha =
      sha;

  }


  /* =====================================
     UPLOAD
  ====================================== */

  const response =
    UrlFetchApp.fetch(
      apiUrl,
      {
        method: 'put',
        headers: headers,
        contentType:
          'application/json',
        payload:
          JSON.stringify(payload),
        muteHttpExceptions:
          true
      }
    );


  const code =
    response.getResponseCode();


  if (
    code !== 200 &&
    code !== 201
  ) {

    throw new Error(
      'GitHub upload failed: ' +
      response.getContentText()
    );

  }


  return {

    fileName:
      fileName,

    path:
      path,

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
   UPLOAD IMAGE FROM AUTOMATIC URL
========================================= */

function uploadVocabularyImageFromUrl(
  imageUrl,
  word,
  category
) {

  if (
    !imageUrl ||
    !String(imageUrl).trim()
  ) {

    throw new Error(
      'Automatic image URL is missing.'
    );

  }


  const response =
    UrlFetchApp.fetch(
      imageUrl,
      {
        method: 'get',
        followRedirects: true,
        muteHttpExceptions: true
      }
    );


  const code =
    response.getResponseCode();


  if (
    code < 200 ||
    code >= 300
  ) {

    throw new Error(
      'Unable to download automatic image. HTTP ' +
      code
    );

  }


  const blob =
    response.getBlob();


  let mimeType =
    blob.getContentType();


  if (
    !mimeType ||
    mimeType.indexOf('image/') !== 0
  ) {

    mimeType =
      'image/jpeg';

  }


  const bytes =
    blob.getBytes();


  if (
    !bytes ||
    bytes.length === 0
  ) {

    throw new Error(
      'Automatic image is empty.'
    );

  }


  const base64 =
    Utilities.base64Encode(bytes);


  return uploadVocabularyImage(
    {
      base64: base64,
      mimeType: mimeType
    },
    word,
    category
  );

}


/* =========================================
   IMAGE EXTENSION
========================================= */

function getImageExtension_(mimeType) {

  const map = {

    'image/jpeg':
      '.jpg',

    'image/jpg':
      '.jpg',

    'image/png':
      '.png',

    'image/webp':
      '.webp',

    'image/gif':
      '.gif'

  };


  return (
    map[mimeType] ||
    '.jpg'
  );

}


/* =========================================
   AUTOMATIC WORD DATA
   Merriam-Webster + Wiktionary fallback
========================================= */

function getAutomaticWordData(word) {

  word = String(word || '').trim();

  if (!word) {
    throw new Error('Please enter a word.');
  }

  const cleanWord = word.toLowerCase();

  let meaning = '';
  let partOfSpeech = '';
  let ipa = '';

  /* =====================================
     MERRIAM-WEBSTER — LEARNER'S DICTIONARY
     Settings!A1
  ===================================== */

  const mwLearner = getMerriamWebsterEntry_(
    cleanWord,
    'learners',
    'A1'
  );

  if (mwLearner) {

    meaning = mwLearner.meaning;
    partOfSpeech = mwLearner.partOfSpeech;
    ipa = mwLearner.ipa;

  }


  /* =====================================
     MERRIAM-WEBSTER — ELEMENTARY
     Settings!A2
  ===================================== */

  if (!meaning) {

    const mwElementary =
      getMerriamWebsterEntry_(
        cleanWord,
        'sd2',
        'A2'
      );

    if (mwElementary) {

      meaning =
        mwElementary.meaning;

      partOfSpeech =
        mwElementary.partOfSpeech;

      ipa =
        mwElementary.ipa;

    }

  }


  /* =====================================
     FINAL CHECK
  ===================================== */

  if (!meaning) {

    throw new Error(
      'Unable to find dictionary information for "' +
      word +
      '".'
    );

  }


  /* =====================================
     PRONUNCIATION OVERRIDES
  ===================================== */

  const ipaOverrides = {

    tomato: '/təˈmeɪtoʊ/',
    potato: '/pəˈteɪtoʊ/',
    eggplant: '/ˈɛɡ.plænt/',
    carrot: '/ˈkærət/',
    cabbage: '/ˈkæbɪdʒ/',
    helicopter: '/ˈhɛlɪˌkɑːptər/',
    rocket: '/ˈrɑːkɪt/',
    onion: '/ˈʌnjən/',
    train: '/treɪn/',
    jeep: '/dʒiːp/',
    engine: '/ˈɛndʒɪn/',
    firetruck: '/ˈfaɪərtrʌk/',
    while: '/waɪl/',
    some: '/sʌm/',
    how: '/haʊ/',
    soon: '/suːn/',
    behind: '/bɪˈhaɪnd/',
    everyone: '/ˈɛvriwʌn/',
    that: '/ðæt/',
    which: '/wɪtʃ/',
    together: '/təˈɡɛðər/',
    king: '/kɪŋ/'

  };


  if (ipaOverrides[cleanWord]) {

    ipa =
      ipaOverrides[cleanWord];

  }


  /* =====================================
     PHONICS
  ===================================== */

  const phonics =
    generatePhonics_(word);


  /* =====================================
     SYLLABLE
  ===================================== */

  const syllable =
    countSyllables_(word);


  /* =====================================
     CATEGORY
  ===================================== */

  const category =
    determineCategoryAutomatically_(
      meaning,
      partOfSpeech,
      word
    );


  /* =====================================
     LEVEL
  ===================================== */

  const level =
    determineLevelAutomatically_(
      word,
      meaning,
      partOfSpeech,
      syllable
    );


  /* =====================================
     RETURN
  ===================================== */

  return {

    Word: word,

    Meaning: meaning,

    Phonics: phonics,

    IPA: ipa,

    Syllable: syllable,

    Category: category,

    Level: level

  };

}


/* =========================================
   MERRIAM-WEBSTER LOOKUP
========================================= */

function getMerriamWebsterEntry_(
  word,
  dictionary,
  keyCell
) {

  try {

    const settings =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Settings');


    if (!settings) {

      console.log(
        'Settings sheet not found.'
      );

      return null;

    }


    const apiKey =
      String(
        settings
          .getRange(keyCell)
          .getDisplayValue()
      ).trim();


    if (!apiKey) {

      console.log(
        'No Merriam-Webster key in ' +
        keyCell
      );

      return null;

    }


    const url =
      'https://www.dictionaryapi.com/api/v3/references/' +
      dictionary +
      '/json/' +
      encodeURIComponent(word) +
      '?key=' +
      encodeURIComponent(apiKey);


    const response =
      UrlFetchApp.fetch(
        url,
        {
          method: 'get',
          muteHttpExceptions: true
        }
      );


    const code =
      response.getResponseCode();


    console.log(
      'Merriam-Webster ' +
      dictionary +
      ' response: ' +
      code
    );


    if (code !== 200) {

      return null;

    }


    const data =
      JSON.parse(
        response.getContentText()
      );


    /*
     * Merriam-Webster returns an array
     * containing dictionary entries.
     *
     * If there is no exact entry it can
     * return an array of suggestions.
     */

    if (
      !Array.isArray(data) ||
      !data.length
    ) {

      return null;

    }


    /*
     * Look for a real dictionary entry.
     */

    let entry = null;


    for (
      let i = 0;
      i < data.length;
      i++
    ) {

      if (
        data[i] &&
        data[i].def
      ) {

        entry =
          data[i];

        break;

      }

    }


    if (!entry) {

      return null;

    }


    const meaning =
      extractMerriamDefinition_(
        entry
      );


    if (!meaning) {

      return null;

    }


    const partOfSpeech =
      String(
        entry.fl || ''
      ).trim();


    const ipa =
      extractMerriamPronunciation_(
        entry
      );


    return {

      meaning: meaning,

      partOfSpeech:
        partOfSpeech,

      ipa: ipa

    };

  } catch (error) {

    console.log(
      'Merriam-Webster error: ' +
      error
    );

    return null;

  }

}


/* =========================================
   EXTRACT MERRIAM-WEBSTER DEFINITION

   If multiple dictionary meanings exist,
   choose ONE complete, clear, short meaning.
========================================= */

function extractMerriamDefinition_(entry) {

  try {

    if (
      !entry ||
      !Array.isArray(entry.shortdef) ||
      !entry.shortdef.length
    ) {

      return '';

    }


    const definitions =
      entry.shortdef
        .map(function(definition) {

          return String(
            definition || ''
          )
            .replace(/\{bc\}/g, '')
            .replace(/\{it\}/g, '')
            .replace(/\{\/it\}/g, '')
            .replace(/\{b\}/g, '')
            .replace(/\{\/b\}/g, '')
            .replace(/\{ldquo\}/g, '"')
            .replace(/\{rdquo\}/g, '"')
            .replace(/\{rs\}/g, "'")
            .replace(/\{\/rs\}/g, "'")
            .replace(/\s+/g, ' ')
            .trim();

        })
        .filter(function(definition) {

          return definition !== '';

        });


    if (!definitions.length) {
      return '';
    }


    /* =====================================
       ONE COMPLETE DICTIONARY MEANING

       Choose the shortest complete definition.
       We are NOT cutting or rewriting it.
    ====================================== */

    definitions.sort(
      function(a, b) {

        return (
          a.length -
          b.length
        );

      }
    );


    return definitions[0];


  } catch (error) {

    console.log(
      'Definition extraction error: ' +
      error
    );

    return '';

  }

}

/* =========================================
   EXTRACT MERRIAM-WEBSTER PRONUNCIATION
========================================= */

function extractMerriamPronunciation_(entry) {

  try {

    if (!entry) {
      return '';
    }


    /* =====================================
       METHOD 1 — hwi.prs[].mw
    ===================================== */

    if (
      entry.hwi &&
      Array.isArray(entry.hwi.prs)
    ) {

      for (
        let i = 0;
        i < entry.hwi.prs.length;
        i++
      ) {

        const pronunciation =
          entry.hwi.prs[i];


        if (
          pronunciation &&
          pronunciation.mw
        ) {

          return String(
            pronunciation.mw
          ).trim();

        }

      }

    }


    /* =====================================
       METHOD 2 — hwi.prs[].ipa
    ===================================== */

    if (
      entry.hwi &&
      Array.isArray(entry.hwi.prs)
    ) {

      for (
        let i = 0;
        i < entry.hwi.prs.length;
        i++
      ) {

        const pronunciation =
          entry.hwi.prs[i];


        if (
          pronunciation &&
          pronunciation.ipa
        ) {

          return String(
            pronunciation.ipa
          ).trim();

        }

      }

    }


    /* =====================================
       METHOD 3 — search pronunciation
       objects recursively
    ===================================== */

    function findPronunciation(value) {

      if (!value) {
        return '';
      }


      if (typeof value === 'object') {

        if (
          value.ipa &&
          typeof value.ipa === 'string'
        ) {

          return value.ipa.trim();

        }


        if (
          value.mw &&
          typeof value.mw === 'string'
        ) {

          return value.mw.trim();

        }


        if (Array.isArray(value)) {

          for (
            let i = 0;
            i < value.length;
            i++
          ) {

            const result =
              findPronunciation(
                value[i]
              );


            if (result) {
              return result;
            }

          }

        } else {

          for (
            const key in value
          ) {

            const result =
              findPronunciation(
                value[key]
              );


            if (result) {
              return result;
            }

          }

        }

      }


      return '';

    }


    return findPronunciation(entry);

  } catch (error) {

    console.log(
      'Pronunciation extraction error: ' +
      error
    );

    return '';

  }

}

/* =========================================
   AUTOMATIC CATEGORY
========================================= */

function determineCategoryAutomatically_(
  meaning,
  partOfSpeech,
  word
) {

  const text =
    (
      String(word || '') +
      ' ' +
      String(meaning || '')
    ).toLowerCase();

  const pos =
    String(partOfSpeech || '')
      .toLowerCase();


  /* =====================================
     ANIMALS
  ====================================== */

  if (
    /\b(animal|insect|bird|mammal|reptile|fish|creature|organism|pet)\b/
      .test(text)
  ) {

    return 'Animals';

  }


  /* =====================================
     FRUITS
  ====================================== */

  if (
    /\b(fruit|berry|citrus)\b/
      .test(text)
  ) {

    return 'Fruits';

  }


  /* =====================================
     FOOD
  ====================================== */

  if (
    /\b(food|vegetable|dish|meal|ingredient|edible)\b/
      .test(text)
  ) {

    return 'Food';

  }


  /* =====================================
     BODY PARTS
  ====================================== */

  if (
    /\b(body|organ|limb|anatomy|bodily)\b/
      .test(text)
  ) {

    return 'Body Parts';

  }


  /* =====================================
     CLOTHES
  ====================================== */

  if (
    /\b(clothing|clothes|garment|footwear|worn|wear)\b/
      .test(text)
  ) {

    return 'Clothes';

  }


  /* =====================================
     TRANSPORT
  ====================================== */

  if (
    /\b(vehicle|automobile|aircraft|airplane|transport|transportation|travel)\b/
      .test(text)
  ) {

    return 'Transport';

  }


  /* =====================================
     FEELINGS
  ====================================== */

  if (
    /\b(feeling|emotion|emotional|mood|happiness|sadness|anger|fear|joy|excitement)\b/
      .test(text)
  ) {

    return 'Feelings';

  }


  /* =====================================
     GEOGRAPHY
  ====================================== */

  if (
    /\b(place|location|region|country|city|continent|mountain|river|lake|ocean|geographical)\b/
      .test(text)
  ) {

    return 'Geography';

  }


  /* =====================================
     EVENTS
  ====================================== */

  if (
    /\b(event|occasion|happening|activity|celebration)\b/
      .test(text)
  ) {

    return 'Events';

  }


  /* =====================================
     ACTIONS
  ====================================== */

  if (
    pos === 'verb' ||
    pos === 'phrasal verb'
  ) {

    return 'Actions';

  }


  /* =====================================
     DESCRIBING WORDS
  ====================================== */

  if (
    pos === 'adjective' ||
    pos === 'adverb'
  ) {

    return 'Describing Words';

  }


  /* =====================================
     DEFAULT
  ====================================== */

  return 'General';

}

/* =========================================
   AUTOMATIC LEVEL
========================================= */

function determineLevelAutomatically_(
  word,
  meaning,
  partOfSpeech,
  syllable
) {

  const cleanWord =
    String(word || '')
      .trim();


  const cleanMeaning =
    String(meaning || '')
      .trim();


  const syllableCount =
    parseInt(
      String(syllable || '')
        .replace(/\D/g, ''),
      10
    ) || 1;


  let score = 0;


  /* =====================================
     WORD LENGTH
  ====================================== */

  if (
    cleanWord.length >= 10
  ) {

    score += 2;

  } else if (
    cleanWord.length >= 7
  ) {

    score += 1;

  }


  /* =====================================
     SYLLABLES
  ====================================== */

  if (
    syllableCount >= 4
  ) {

    score += 2;

  } else if (
    syllableCount >= 3
  ) {

    score += 1;

  }


  /* =====================================
     DEFINITION COMPLEXITY
  ====================================== */

  if (
    cleanMeaning.length >= 120
  ) {

    score += 2;

  } else if (
    cleanMeaning.length >= 70
  ) {

    score += 1;

  }


  /* =====================================
   ABSTRACT / TECHNICAL SIGNALS
====================================== */

if (
  /\b(process|system|theory|principle|method|concept|condition|property|structure|scientific|technical|formal|political|economic|philosophical)\b/i
    .test(cleanMeaning)
) {

  score += 2;

}

  /* =====================================
     SIMPLE COMMON VERBS
  ====================================== */

  const pos =
    String(partOfSpeech || '')
      .toLowerCase();


  if (
    pos === 'verb' &&
    syllableCount <= 2 &&
    cleanWord.length <= 6
  ) {

    score -= 1;

  }


  /* =====================================
     FINAL LEVEL
  ====================================== */

  if (
    score <= 1
  ) {

    return 'Basic';

  }


  if (
    score <= 4
  ) {

    return 'Intermediate';

  }


  return 'Advanced';

}


/* =========================================
   PHONICS
========================================= */

function generatePhonics_(word) {

  word =
    String(word || '')
      .toLowerCase()
      .replace(
        /[^a-z]/g,
        ''
      );


  if (!word) {
    return '';
  }


  const groups = [

    'tion',
    'sion',
    'ture',
    'ough',
    'eigh',
    'igh',
    'ph',
    'sh',
    'ch',
    'th',
    'wh',
    'ck',
    'ng',
    'qu',
    'oo',
    'ee',
    'ea',
    'ai',
    'ay',
    'oa',
    'ow',
    'ou',
    'oi',
    'oy',
    'ar',
    'er',
    'ir',
    'ur'

  ];


  const result = [];

  let i = 0;


  while (
    i < word.length
  ) {

    let found = '';


    for (
      let j = 0;
      j < groups.length;
      j++
    ) {

      const group =
        groups[j];


      if (
        word.substring(
          i,
          i + group.length
        ) === group
      ) {

        found =
          group;

        break;

      }

    }


    if (found) {

      result.push(found);

      i += found.length;

    } else {

      result.push(
        word.charAt(i)
      );

      i++;

    }

  }


  return result.join('-');

}


/* =========================================
   SYLLABLE COUNT
========================================= */

function countSyllables_(word) {

  word =
    String(word || '')
      .toLowerCase()
      .replace(
        /[^a-z]/g,
        ''
      );


  if (!word) {
    return '';
  }


  const exceptions = {

    potato: 3,

    tomato: 3,

    eggplant: 2,

    carrot: 2,

    cabbage: 2,

    helicopter: 4,

    rocket: 2,

    vegetable: 4,

    beautiful: 3,

    family: 3,

    every: 3,

    everyone: 4

  };


  if (
    exceptions[word]
  ) {

    return exceptions[word];

  }


  let count = 0;


  const vowelGroups =
    word.match(
      /[aeiouy]+/g
    );


  if (vowelGroups) {

    count =
      vowelGroups.length;

  }


  /* =====================================
     SILENT FINAL E
  ====================================== */

  if (
    word.endsWith('e') &&
    count > 1 &&
    !word.endsWith('le')
  ) {

    count--;

  }


  /* =====================================
     CONSONANT + LE
  ====================================== */

  if (
    word.endsWith('le') &&
    word.length > 2 &&
    !/[aeiou]le$/.test(word)
  ) {

    count++;

  }


  return Math.max(
    1,
    count
  );

}


/* =========================================
   CHECK DUPLICATE WORD
========================================= */

function checkDuplicateWord(word) {

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
    sheet
      .getDataRange()
      .getDisplayValues();


  if (
    data.length <= 1
  ) {

    return false;

  }


  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim();

    });


  const wordColumn =
    headers.indexOf('Word');


  if (
    wordColumn === -1
  ) {

    throw new Error(
      'Word column not found.'
    );

  }


  const searchWord =
    String(word || '')
      .trim()
      .toLowerCase();


  if (!searchWord) {
    return false;
  }


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const existingWord =
      String(
        data[i][wordColumn] || ''
      )
        .trim()
        .toLowerCase();


    if (
      existingWord === searchWord
    ) {

      return true;

    }

  }


  return false;

}

/* =========================================
   FREE IMAGE SEARCH
   WIKIMEDIA + OPENVERSE
========================================= */

function searchFreeImages(word, meaning, category) {

  word = String(word || '').trim();
  meaning = String(meaning || '').trim();
  category = String(category || '').trim();

  if (!word) {
    return {
      images: []
    };
  }

  const results = [];
  const seen = {};

  function addResults(items) {

    if (!Array.isArray(items)) {
      return;
    }

    items.forEach(function(item) {

      if (!item) {
        return;
      }

      const imageUrl =
        String(
          item.imageUrl ||
          item.thumbnailUrl ||
          item.url ||
          ''
        ).trim();

      if (!imageUrl) {
        return;
      }

      if (seen[imageUrl]) {
        return;
      }

      seen[imageUrl] = true;

      results.push({
        imageUrl: imageUrl,
        title:
          String(
            item.title ||
            word
          )
      });

    });

  }


  /* =====================================
     1. WIKIMEDIA EXACT WORD
  ====================================== */

  try {

    addResults(
      searchWikimediaImages_(word)
    );

  } catch (error) {

    console.log(
      'Wikimedia exact search error: ' +
      error
    );

  }


  /* =====================================
     2. OPENVERSE EXACT WORD
  ====================================== */

  if (results.length < 10) {

    try {

      addResults(
        searchOpenverseImages_(word)
      );

    } catch (error) {

      console.log(
        'Openverse exact search error: ' +
        error
      );

    }

  }


  /* =====================================
     3. WORD + CATEGORY
  ====================================== */

  if (
    results.length < 5 &&
    category
  ) {

    const query =
      word +
      ' ' +
      category;


    try {

      addResults(
        searchWikimediaImages_(query)
      );

    } catch (error) {

      console.log(
        'Wikimedia category search error: ' +
        error
      );

    }


    if (results.length < 5) {

      try {

        addResults(
          searchOpenverseImages_(query)
        );

      } catch (error) {

        console.log(
          'Openverse category search error: ' +
          error
        );

      }

    }

  }


  /* =====================================
     4. WORD + IMPORTANT MEANING WORDS
  ====================================== */

  if (
    results.length < 5 &&
    meaning
  ) {

    const meaningWords =
      meaning
        .toLowerCase()
        .replace(
          /[^a-z\s-]/g,
          ' '
        )
        .split(/\s+/)
        .filter(function(item) {

          return (
            item.length >= 4 &&
            !isCommonImageStopWord_(item)
          );

        })
        .slice(0, 3);


    if (meaningWords.length) {

      const query =
        word +
        ' ' +
        meaningWords.join(' ');


      try {

        addResults(
          searchWikimediaImages_(query)
        );

      } catch (error) {

        console.log(
          'Wikimedia meaning search error: ' +
          error
        );

      }


      if (results.length < 5) {

        try {

          addResults(
            searchOpenverseImages_(query)
          );

        } catch (error) {

          console.log(
            'Openverse meaning search error: ' +
            error
          );

        }

      }

    }

  }


  /* =====================================
     5. SPECIAL NAMES
  ====================================== */

  const alternatives = {

    eggplant: [
      'aubergine',
      'brinjal'
    ],

    aubergine: [
      'eggplant',
      'brinjal'
    ],

    brinjal: [
      'eggplant',
      'aubergine'
    ]

  };


  const alternateWords =
    alternatives[
      word.toLowerCase()
    ] || [];


  for (
    let i = 0;
    i < alternateWords.length;
    i++
  ) {

    if (results.length >= 5) {
      break;
    }

    try {

      addResults(
        searchWikimediaImages_(
          alternateWords[i]
        )
      );

    } catch (error) {

      console.log(
        'Alternative Wikimedia search error: ' +
        error
      );

    }

  }


  console.log(
    'Image search for "' +
    word +
    '" returned ' +
    results.length +
    ' images.'
  );


  return {

    images:
      results.slice(0, 5)

  };

}


/* =========================================
   OPENVERSE SEARCH
========================================= */

function searchOpenverseImages_(searchTerm) {

  const url =
    'https://api.openverse.org/v1/images/?q=' +
    encodeURIComponent(
      String(searchTerm || '').trim()
    ) +
    '&page_size=20';


  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Accept: 'application/json'
        }
      }
    );


  const code =
    response.getResponseCode();


  if (code !== 200) {

    console.log(
      'Openverse HTTP ' +
      code +
      ': ' +
      response.getContentText()
    );

    return [];

  }


  let data = null;


  try {

    data =
      JSON.parse(
        response.getContentText()
      );

  } catch (error) {

    console.log(
      'Openverse JSON error: ' +
      error
    );

    return [];

  }


  if (
    !data ||
    !Array.isArray(
      data.results
    )
  ) {

    return [];

  }


  return data.results
    .map(function(item) {

      if (!item) {
        return null;
      }

      return {

        imageUrl:
          item.thumbnail ||
          item.url ||
          '',

        title:
          item.title ||
          '',

        description:
          item.description ||
          '',

        tags:
          Array.isArray(item.tags)
            ? item.tags
                .map(function(tag) {

                  if (
                    typeof tag === 'string'
                  ) {

                    return tag;

                  }

                  return (
                    tag.name ||
                    ''
                  );

                })
                .join(' ')
            : ''

      };

    })
    .filter(function(item) {

      return (
        item &&
        item.imageUrl
      );

    });

}


/* =========================================
   WIKIMEDIA COMMONS SEARCH
========================================= */

function searchWikimediaImages_(searchTerm) {

  const query =
    String(searchTerm || '').trim();


  if (!query) {
    return [];
  }


  const url =
    'https://commons.wikimedia.org/w/api.php' +
    '?action=query' +
    '&format=json' +
    '&generator=search' +
    '&gsrsearch=' +
    encodeURIComponent(query) +
    '&gsrnamespace=6' +
    '&gsrlimit=20' +
    '&prop=imageinfo' +
    '&iiprop=url|extmetadata' +
    '&iiurlwidth=800';


  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Accept: 'application/json'
        }
      }
    );


  const code =
    response.getResponseCode();


  if (code !== 200) {

    console.log(
      'Wikimedia HTTP ' +
      code +
      ': ' +
      response.getContentText()
    );

    return [];

  }


  let data = null;


  try {

    data =
      JSON.parse(
        response.getContentText()
      );

  } catch (error) {

    console.log(
      'Wikimedia JSON error: ' +
      error
    );

    return [];

  }


  if (
    !data ||
    !data.query ||
    !data.query.pages
  ) {

    return [];

  }


  const pages =
    data.query.pages;


  return Object.keys(pages)
    .map(function(id) {

      const page =
        pages[id];


      if (!page) {
        return null;
      }


      const info =
        page.imageinfo &&
        page.imageinfo[0];


      if (!info) {
        return null;
      }


      const metadata =
        info.extmetadata || {};


      let description = '';


      if (
        metadata.ImageDescription &&
        metadata.ImageDescription.value
      ) {

        description =
          String(
            metadata.ImageDescription.value
          );

      }


      return {

        imageUrl:
          info.thumburl ||
          info.url ||
          '',

        title:
          page.title
            ? String(
                page.title
              ).replace(
                /^File:/i,
                ''
              )
            : '',

        description:
          description,

        tags: ''

      };

    })
    .filter(function(item) {

      return (
        item &&
        item.imageUrl
      );

    });

}


/* =========================================
   COMMON IMAGE SEARCH STOP WORDS
========================================= */

function isCommonImageStopWord_(word) {

  const stopWords = [

    'this',
    'that',
    'these',
    'those',
    'also',
    'called',
    'used',
    'with',
    'from',
    'into',
    'than',
    'then',
    'they',
    'them',
    'their',
    'there',
    'where',
    'which',
    'when',
    'what',
    'have',
    'has',
    'been',
    'being',
    'very',
    'more',
    'most',
    'some',
    'such',
    'word',
    'means',
    'meaning',
    'thing',
    'things',
    'type',
    'kind'

  ];


  return stopWords.includes(
    String(word || '')
      .toLowerCase()
  );

}

/* =========================================
   ADD IMAGE RESULT
========================================= */

function addImageResult_(
  images,
  item,
  word,
  source
) {

  if (!item) {
    return;
  }


  const imageUrl =
    item.url ||
    item.imageUrl ||
    item.image ||
    item.src ||
    '';


  if (!imageUrl) {
    return;
  }


  const exists =
    images.some(
      function(existing) {

        return (
          existing.imageUrl ===
          imageUrl
        );

      }
    );


  if (exists) {
    return;
  }


  images.push({

    imageUrl:
      String(imageUrl),

    thumbnailUrl:
      String(
        item.thumbnail ||
        item.thumbnailUrl ||
        imageUrl
      ),

    title:
      String(
        item.description ||
        item.title ||
        word
      ),

    source:
      source

  });

}