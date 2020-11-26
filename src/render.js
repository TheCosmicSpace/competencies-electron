const xlsxParser = require('xls-parser')

// Global vars
const form = document.querySelector('.main-form')
let file, searchLesson, competencies 

// Init vars
const init = () => {
  file = null
  searchLesson = ''
  competencies = []
}
init()

// Select file event 
const onSelectFile = (event) => {
  file = event.target.files ? event.target.files[0] : null
  console.log(file)
  if (!file) return alert('Выберите файл')
  document.querySelector('.file-name').innerHTML = file.name
}
document.querySelector('.main-file-input').addEventListener('change', onSelectFile)

const enterSearchLesson = () => {
  const { value } = document.querySelector('.lesson-name')
  console.log(value)
  searchLesson = value
}

const constructCompetencies = collection => {
  return collection.reduce((acc, lesson) => {
    const lessonName = lesson['Содержание'] || lesson['Наименование'] || ''
    console.log(lessonName);
    if (lessonName.toLowerCase().trim() === searchLesson.toLowerCase().trim() &&
        lesson['Формируемые компетенции']) acc.push(lesson['Формируемые компетенции'])
    return acc
  }, [])
}

const generateWordDocument = competencies => {
  if (!competencies.length) return alert('Ничего не найдено')
  const doc = new docx.Document()
  const children = competencies.map(item => {
    return new docx.TextRun(item)
  })
  doc.addSection({
    properties: {},
    children: [
      new docx.Paragraph({
        children
      })
    ]
  })
  saveDocumentToFile(doc, searchLesson)
}

// Save Word Document
const saveDocumentToFile = (doc, fileName) => {
  docx.Packer.toBlob(doc).then(blob => {
    saveAs(blob, fileName)
  })
}

// Check Sheets Lists 
const checkSheet = parsedData => {
  let competencies = []
  for (const [key, value] of Object.entries(parsedData)) {
    if (key.match(/компетенции/gi)) {
      competencies = [...competencies, ...new Set(constructCompetencies(value))]
    }
  }
  generateWordDocument(competencies)
}

// Parse File
const parseSelectedFile = () => {
  xlsxParser
    .onFileSelection(file)
    .then(data => checkSheet(data))
}

// Event submit form
form.addEventListener('submit', e => {
  e.preventDefault()
  
  // Check needed
  enterSearchLesson()
  if (!file) return alert('Необходимо выбрать фаил')
  if (!searchLesson) return alert('Необходимо ввести название предмета')
  
  // Start Parse
  parseSelectedFile()
})
