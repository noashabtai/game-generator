require('dotenv').config();
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// שמירת שם הקובץ האחרון
let lastFilename = null;

// ==================== ניהול סיסמאות ====================

const ADMIN_PASSWORD = 'Noa712';

const UNIT_PASSWORDS = {
  'ALPHA7821': { name: 'גדוד א׳', used: false },
  'BRAVO5493': { name: 'גדוד ב׳', used: false },
  'CHARLIE2847': { name: 'גדוד ג׳', used: false },
  'DELTA9156': { name: 'גדוד ד׳', used: false },
  'ECHO6734': { name: 'גדוד ה׳', used: false },
  'FOXTROT4521': { name: 'גדוד ו׳', used: false },
  'GOLF8309': { name: 'גדוד ז׳', used: false },
  'HOTEL1675': { name: 'גדוד ח׳', used: false },
  'INDIA7942': { name: 'גדוד ט׳', used: false },
  'JULIET3568': { name: 'גדוד י׳', used: false },
  'KILO8425': { name: 'גדוד יא׳', used: false },
  'LIMA5912': { name: 'גדוד יב׳', used: false }
};

// ==================== עיבוד הנתונים ====================

async function generateMissingWords(categories, existingWords, totalWords = 800) {
  try {
    const missingCount = totalWords - existingWords.length;
    
    if (missingCount <= 0) {
      return existingWords.slice(0, totalWords);
    }

    const categoriesStr = categories.join(', ');
    const wordsStr = existingWords.slice(0, 50).join(', ');

    const prompt = `אתה עוזר ליצור מילים למשחק קופסה גדודי.

הקטגוריות הן: ${categoriesStr}

המילים שכבר יש: ${wordsStr}... (ועוד)

צור לי ${missingCount} מילים חדשות שלא חוזרות על עצמן, קשורות לקטגוריות האלה.
המילים צריכות להיות:
- בעברית
- קצרות (מילה או שתיים)
- רלוונטיות לגדוד/חיים/משפחה
- מגוונות

החזר רק רשימה של מילים, מופרדות בשורה חדשה. ללא מספורים, ללא הסברים.`;

    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });
    const result = await model.generateContent(prompt);
    const generatedText = result.response.text();

    const newWords = generatedText.split('\n')
      .map(w => w.trim())
      .filter(w => w.length > 0 && !existingWords.includes(w))
      .slice(0, missingCount);

    return [...existingWords, ...newWords].slice(0, totalWords);
  } catch (error) {
    console.error('Error generating words:', error.message);
    throw error;
  }
}

async function extractQuestionsFromStories(stories, numQuestions = 48) {
  try {
    const storiesText = stories
      .map(s => `[${s.company}] ${s.story}`)
      .join('\n\n');

    const prompt = `אתה עוזר לחלץ שאלות משחק מסיפורים גדודיים.

הסיפורים:
${storiesText}

צור לי ${numQuestions} שאלות/משימות קצרות בנוסח "ספרו לנו..." או "תסביר לי..." שמתוך הסיפורים האלה.
השאלות צריכות להיות:
- קצרות (עד 10 מילים כל אחת)
- מעוררות סיפור וזיכרון
- קשורות לתוכן של הסיפורים
- מגוונות

החזר רק רשימה של שאלות, מופרדות בשורה חדשה. ללא מספורים.`;

    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });
    const result = await model.generateContent(prompt);
    const generatedText = result.response.text();

    const questions = generatedText.split('\n')
      .map(q => q.trim())
      .filter(q => q.length > 0)
      .slice(0, numQuestions);

    return questions;
  } catch (error) {
    console.error('Error extracting questions:', error.message);
    throw error;
  }
}

// ==================== קריאת קבצים ====================

function readExcelFile(buffer, sheetName = 0) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[sheetName]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    return data.flat().filter(cell => cell && String(cell).trim().length > 0);
  } catch (error) {
    console.error('Error reading Excel:', error.message);
    throw error;
  }
}

function readStoriesFile(buffer) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    return data.map(row => ({
      company: row['פלוגה'] || row['Company'] || 'לא מוגדר',
      story: row['סיפור'] || row['Story'] || row['סיפור/זיכרון'] || ''
    })).filter(s => s.story.trim().length > 0);
  } catch (error) {
    console.error('Error reading stories:', error.message);
    throw error;
  }
}

// ==================== יצוא Excel ====================

function createExcelFile(words, questions, metadata) {
  try {
    const wb = xlsx.utils.book_new();

    // גיליון 1: קלפים בסיסיים
    const basicCards = words.map((word, idx) => ({
      'קלף #': idx + 1,
      'מילה': word,
      'קטגוריה': '(עיצוב בעיצובנית)'
    }));
    const ws1 = xlsx.utils.json_to_sheet(basicCards);
    xlsx.utils.book_append_sheet(wb, ws1, 'קלפים בסיסיים');

    // גיליון 2: קלפי סיפור
    const storyCards = questions.map((q, idx) => ({
      'קלף #': idx + 1,
      'שאלה': q,
      'פלוגה': metadata.companies[idx % metadata.companies.length] || 'מחולק'
    }));
    const ws2 = xlsx.utils.json_to_sheet(storyCards);
    xlsx.utils.book_append_sheet(wb, ws2, 'קלפי סיפור');

    // גיליון 3: מטא-נתונים
    const metaData = [
      ['שם המשחק:', metadata.gameName],
      ['סלוגן:', metadata.slogan],
      ['שם קלפי סיפור:', metadata.storyCardName],
      ['שם קלפי אזרחות:', metadata.civilianCardName],
      ['כמות מילים סה״כ:', words.length],
      ['כמות שאלות:', questions.length],
      ['כמות פלוגות:', metadata.companies.length],
      ['תאריך יצוא:', new Date().toLocaleDateString('he-IL')],
      [''],
      ['קטגוריות (למידע בלבד):', metadata.categories.join(', ')]
    ];
    const ws3 = xlsx.utils.json_to_sheet(metaData, { header: 1 });
    xlsx.utils.book_append_sheet(wb, ws3, 'מטא-נתונים');

    return xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
  } catch (error) {
    console.error('Error creating Excel:', error.message);
    throw error;
  }
}

// ==================== API Routes ====================

// בדיקת סיסמה
app.post('/api/verify-password', (req, res) => {
  try {
    const { password } = req.body;

    if (!password) {
      return res.status(400).json({ error: 'נא הכנסו סיסמה' });
    }

    // בדיקה אם זו סיסמת Admin
    if (password === ADMIN_PASSWORD) {
      return res.json({ 
        valid: true, 
        type: 'admin',
        message: 'ברוכה הבאה Admin!'
      });
    }

    // בדיקה אם זו סיסמה של גדוד
    if (UNIT_PASSWORDS[password]) {
      const unitInfo = UNIT_PASSWORDS[password];
      if (unitInfo.used) {
        return res.status(403).json({ 
          error: 'סיסמה זו כבר שימשה. אנא צרו קשר עם נעה לסיסמה חדשה.' 
        });
      }
      return res.json({ 
        valid: true, 
        type: 'unit',
        unitName: unitInfo.name,
        message: `ברוכים הבאים ${unitInfo.name}!`
      });
    }

    return res.status(401).json({ error: 'סיסמה לא נכונה' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Admin - קבלת רשימת סיסמאות
app.get('/api/admin/passwords', (req, res) => {
  try {
    const adminPassword = req.query.adminPass;
    
    if (adminPassword !== ADMIN_PASSWORD) {
      return res.status(403).json({ error: 'גישה נדחתה' });
    }

    const passwordsList = Object.entries(UNIT_PASSWORDS).map(([pass, info]) => ({
      password: pass,
      unitName: info.name,
      used: info.used
    }));

    res.json({ passwords: passwordsList });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Admin - יצירת סיסמה חדשה
app.post('/api/admin/new-password', (req, res) => {
  try {
    const { adminPass, unitIndex, newPassword } = req.body;

    if (adminPass !== ADMIN_PASSWORD) {
      return res.status(403).json({ error: 'גישה נדחתה' });
    }

    const unitKeys = Object.keys(UNIT_PASSWORDS);
    if (unitIndex < 0 || unitIndex >= unitKeys.length) {
      return res.status(400).json({ error: 'אינדקס גדוד לא תקין' });
    }

    const oldPass = unitKeys[unitIndex];
    const unitName = UNIT_PASSWORDS[oldPass].name;

    // מחק את הסיסמה הישנה
    delete UNIT_PASSWORDS[oldPass];

    // הוסף סיסמה חדשה
    UNIT_PASSWORDS[newPassword] = { name: unitName, used: false };

    res.json({ 
      success: true, 
      message: `סיסמה חדשה ליצור ${unitName}: ${newPassword}` 
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// עיבוד קבצים (חייבת סיסמה תקינה)
app.post('/api/process', upload.fields([
  { name: 'wordsFile', maxCount: 1 },
  { name: 'storiesFile', maxCount: 1 }
]), async (req, res) => {
  try {
    const { password, gameName, slogan, storyCardName, civilianCardName, categories, companies } = req.body;

    // בדיקת סיסמה
    if (!UNIT_PASSWORDS[password] && password !== ADMIN_PASSWORD) {
      return res.status(403).json({ error: 'סיסמה לא תקינה' });
    }

    // סימון סיסמה כשימשה (אם היא סיסמה של גדוד)
    if (UNIT_PASSWORDS[password] && password !== ADMIN_PASSWORD) {
      UNIT_PASSWORDS[password].used = true;
    }

    const categoriesList = JSON.parse(categories || '[]');
    const companiesList = JSON.parse(companies || '[]');

    let words = [];
    let stories = [];

    if (req.files.wordsFile) {
      words = readExcelFile(req.files.wordsFile[0].buffer);
    }

    if (req.files.storiesFile) {
      stories = readStoriesFile(req.files.storiesFile[0].buffer);
    }

    res.json({
      status: 'processing',
      message: `קיבלתי ${words.length} מילים ו-${stories.length} סיפורים. מעבדתי...`,
      progress: 'generating'
    });

    // עיבוד בריקע
    setTimeout(async () => {
      try {
        console.log('🤖 Gemini עובד על המילים...');
        const allWords = await generateMissingWords(categoriesList, words, 800);

        console.log('🤖 Gemini עובד על השאלות...');
        const allQuestions = await extractQuestionsFromStories(stories, 48);

        console.log('📊 יוצר Excel...');
        const excelBuffer = createExcelFile(allWords, allQuestions, {
          gameName,
          slogan,
          storyCardName,
          civilianCardName,
          categories: categoriesList,
          companies: companiesList
        });

        const filename = `game-${Date.now()}.xlsx`;
        fs.writeFileSync(`/tmp/${filename}`, excelBuffer);
        lastFilename = filename;

        console.log(`✅ קובץ יצוא: ${filename}`);
      } catch (error) {
        console.error('Processing error:', error);
      }
    }, 0);

  } catch (error) {
    console.error('API error:', error);
    res.status(500).json({ error: error.message });
  }
});

// הורדת קובץ - גרסה מתוקנת
app.get('/api/download/latest', (req, res) => {
  try {
    if (!lastFilename) {
      return res.status(404).json({ error: 'קובץ לא נמצא. אנא העלו קבצים תחילה.' });
    }

    const filepath = path.join('/tmp', lastFilename);

    if (!fs.existsSync(filepath)) {
      return res.status(404).json({ error: 'קובץ לא זמין. אנא נסו שוב.' });
    }

    res.download(filepath, `game-generator-${Date.now()}.xlsx`, () => {
      try {
        fs.unlinkSync(filepath);
        lastFilename = null;
      } catch (e) {
        console.error('Error deleting file:', e);
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ==================== Start Server ====================

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🎲 משחק קופסה גדודי - מחולל תוכן`);
  console.log(`🌐 שרת פועל על: http://localhost:${PORT}`);
  console.log(`🤖 Gemini API: ${process.env.GEMINI_API_KEY ? '✅ מוגדר' : '❌ לא מוגדר'}`);
  console.log(`🔐 Admin Password: ${ADMIN_PASSWORD}`);
  console.log(`📝 Unit Passwords: ${Object.keys(UNIT_PASSWORDS).length} גדודים`);
});
