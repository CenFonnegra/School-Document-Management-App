document.addEventListener('DOMContentLoaded', () => {
    const newDocButton = document.getElementById('new-doc');
    const newFromTemplateButton = document.getElementById('new-from-template');
    const importFileInput = document.getElementById('import-file');
    const saveDocButton = document.getElementById('save-doc');
    const deleteDocButton = document.getElementById('delete-doc');
    const searchBox = document.getElementById('search-box');
    const editor = document.getElementById('editor');
  
    // Handle new document creation
    newDocButton.addEventListener('click', () => {
      editor.value = '';
      editor.placeholder = 'Start writing your document...';
    });
  
    // Handle new document from template creation
    newFromTemplateButton.addEventListener('click', () => {
      editor.value = 'Template Content: \n\n';
    });
  
    // Handle file import
    importFileInput.addEventListener('change', async (event) => {
      const file = event.target.files[0];
      if (file) {
        const fileType = file.type;
  
        if (fileType === 'application/pdf') {
          // Handle PDF files
          const pdfBytes = await file.arrayBuffer();
          const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
          const textContent = await extractPdfText(pdfDoc);
          editor.value = textContent;
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
          // Handle Word files
          const reader = new FileReader();
          reader.onload = async (e) => {
            const arrayBuffer = e.target.result;
            const doc = await docx.Packer.toDocument(new Uint8Array(arrayBuffer));
            const textContent = await extractDocxText(doc);
            editor.value = textContent;
          };
          reader.readAsArrayBuffer(file);
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
          // Handle Excel files
          const reader = new FileReader();
          reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const textContent = extractXlsxText(workbook);
            editor.value = textContent;
          };
          reader.readAsArrayBuffer(file);
        } else {
          // Handle other text files
          const reader = new FileReader();
          reader.onload = (e) => {
            editor.value = e.target.result;
          };
          reader.readAsText(file);
        }
      }
    });
  
    // Extract text from PDF document
    async function extractPdfText(pdfDoc) {
      const pages = pdfDoc.getPages();
      let textContent = '';
      for (const page of pages) {
        const text = await page.getTextContent();
        textContent += text.items.map((item) => item.str).join(' ') + '\n';
      }
      return textContent;
    }
  
    // Extract text from Word document
    async function extractDocxText(doc) {
      const paragraphs = doc.paragraphs;
      let textContent = '';
      for (const paragraph of paragraphs) {
        textContent += paragraph.children.map((child) => child.text).join('') + '\n';
      }
      return textContent;
    }
  
    // Extract text from Excel document
    function extractXlsxText(workbook) {
      const sheetNames = workbook.SheetNames;
      let textContent = '';
      sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetText = XLSX.utils.sheet_to_txt(worksheet);
        textContent += sheetText + '\n';
      });
      return textContent;
    }
  
    // Handle save document
    saveDocButton.addEventListener('click', () => {
      const blob = new Blob([editor.value], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.txt';
      a.click();
      URL.revokeObjectURL(url);
    });
  
    // Handle delete document
    deleteDocButton.addEventListener('click', () => {
      const confirmed = confirm('Are you sure you want to delete this document?');
      if (confirmed) {
        editor.value = '';
        editor.placeholder = 'Start writing your document...';
      }
    });
  
    // Handle search functionality
    searchBox.addEventListener('input', () => {
      const searchText = searchBox.value.toLowerCase();
      const editorText = editor.value.toLowerCase();
      const index = editorText.indexOf(searchText);
  
      if (index !== -1) {
        editor.focus();
        editor.setSelectionRange(index, index + searchText.length);
      }
    });
  });
  