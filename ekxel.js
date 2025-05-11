document.addEventListener('DOMContentLoaded', function () {
   const spreadsheet = document.getElementById('spreadsheet');
   const headerRow = document.getElementById('headerRow');
   const dataBody = document.getElementById('dataBody');
   const csvFileInput = document.getElementById('csvFile');
   const exportBtn = document.getElementById('exportBtn');
   const addColBeforeBtn = document.getElementById('addColBeforeBtn');
   const addColAfterBtn = document.getElementById('addColAfterBtn');
   const removeColBtn = document.getElementById('removeColBtn');
   const moveColLeftBtn = document.getElementById('moveColLeftBtn');
   const moveColRightBtn = document.getElementById('moveColRightBtn');
   const addRowBeforeBtn = document.getElementById('addRowBeforeBtn');
   const addRowAfterBtn = document.getElementById('addRowAfterBtn');
   const removeRowBtn = document.getElementById('removeRowBtn');
   const moveRowUpBtn = document.getElementById('moveRowUpBtn');
   const moveRowDownBtn = document.getElementById('moveRowDownBtn');

   let activeCell = null;
   let data = [];
   let headers = [];
   let sortState = {}; // Track sorting state for each column

   // Initialize with empty data
   initializeSpreadsheet(3, 3);

   // Event listeners
   csvFileInput.addEventListener('change', handleFileUpload);
   exportBtn.addEventListener('click', exportToCSV);
   addColBeforeBtn.addEventListener('click', () => addColumn('before'));
   addColAfterBtn.addEventListener('click', () => addColumn('after'));
   removeColBtn.addEventListener('click', removeColumn);
   moveColLeftBtn.addEventListener('click', () => moveColumn('left'));
   moveColRightBtn.addEventListener('click', () => moveColumn('right'));
   addRowBeforeBtn.addEventListener('click', () => addRow('before'));
   addRowAfterBtn.addEventListener('click', () => addRow('after'));
   removeRowBtn.addEventListener('click', removeRow);
   moveRowUpBtn.addEventListener('click', () => moveRow('up'));
   moveRowDownBtn.addEventListener('click', () => moveRow('down'));

   function initializeSpreadsheet(cols, rows) {
      headers = Array(cols).fill(0).map((_, i) => `Column ${i+1}`);
      data = Array(rows).fill(0).map(() => Array(cols).fill(''));
      renderSpreadsheet();
   }

   function renderSpreadsheet() {
      // Clear existing content
      headerRow.innerHTML = '';
      dataBody.innerHTML = '';

      // Render headers
      headers.forEach((header, colIndex) => {
         const th = document.createElement('th');
         th.textContent = header;
         th.addEventListener('click', () => sortColumn(colIndex));

         // Add sort indicator if sorted
         if (sortState[colIndex]) {
            th.classList.add(sortState[colIndex] === 'asc' ? 'sort-asc' : 'sort-desc');
         }

         headerRow.appendChild(th);
      });

      // Render data rows
      data.forEach((row, rowIndex) => {
         const tr = document.createElement('tr');

         row.forEach((cell, colIndex) => {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.value = cell;

            input.addEventListener('focus', () => {
               if (activeCell) {
                  activeCell.classList.remove('active-cell');
               }
               activeCell = td;
               td.classList.add('active-cell');
            });

            input.addEventListener('blur', () => {
               data[rowIndex][colIndex] = input.value;
            });

            input.addEventListener('keydown', (e) => {
               if (e.key === 'Enter') {
                  e.preventDefault();
                  const nextRow = tr.nextElementSibling;
                  if (nextRow) {
                     nextRow.children[colIndex].querySelector('input').focus();
                  }
               } else if (e.key === 'Tab') {
                  e.preventDefault();
                  const nextCell = e.shiftKey ?
                     td.previousElementSibling || tr.previousElementSibling?.lastElementChild :
                     td.nextElementSibling || tr.nextElementSibling?.firstElementChild;
                  if (nextCell) {
                     nextCell.querySelector('input').focus();
                  }
               }
            });

            td.appendChild(input);
            tr.appendChild(td);
         });

         dataBody.appendChild(tr);
      });
   }

   function handleFileUpload(event) {
      const file = event.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (e) {
         const content = e.target.result;
         parseCSV(content);
      };
      reader.readAsText(file);
   }

   function parseCSV(csvText) {
      const lines = csvText.split('\n').filter(line => line.trim() !== '');
      if (lines.length === 0) return;

      // Parse headers (first line)
      headers = lines[0].split(',').map(h => h.trim());

      // Parse data rows
      data = [];
      for (let i = 1; i < lines.length; i++) {
         const row = lines[i].split(',');
         // Ensure each row has the same number of columns as headers
         const filledRow = Array(headers.length).fill('');
         for (let j = 0; j < Math.min(row.length, headers.length); j++) {
            filledRow[j] = row[j].trim();
         }
         data.push(filledRow);
      }

      // Reset sort state
      sortState = {};
      renderSpreadsheet();
   }

   function exportToCSV() {
      // Prepare CSV content
      let csvContent = headers.join(',') + '\n';
      data.forEach(row => {
         csvContent += row.join(',') + '\n';
      });

      // Create download link
      const blob = new Blob([csvContent], {
         type: 'text/csv;charset=utf-8;'
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', 'spreadsheet.csv');
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
   }

   function addColumn(position) {
      if (!activeCell) return;

      const colIndex = Array.from(activeCell.parentNode.children).indexOf(activeCell);
      const newColIndex = position === 'before' ? colIndex : colIndex + 1;

      // Update headers
      const newHeader = `Column ${headers.length + 1}`;
      headers.splice(newColIndex, 0, newHeader);

      // Update data
      data.forEach(row => {
         row.splice(newColIndex, 0, '');
      });

      // Reset sort state
      sortState = {};
      renderSpreadsheet();
   }

   function removeColumn() {
      if (!activeCell) return;

      const colIndex = Array.from(activeCell.parentNode.children).indexOf(activeCell);
      if (headers.length <= 1) return; // Don't remove the last column

      // Update headers
      headers.splice(colIndex, 1);

      // Update data
      data.forEach(row => {
         row.splice(colIndex, 1);
      });

      // Reset sort state
      sortState = {};
      renderSpreadsheet();
   }

   function moveColumn(direction) {
      if (!activeCell) return;

      const colIndex = Array.from(activeCell.parentNode.children).indexOf(activeCell);
      const newIndex = direction === 'left' ? colIndex - 1 : colIndex + 1;

      // Check if move is possible
      if (newIndex < 0 || newIndex >= headers.length) return;

      // Move header
      const headerToMove = headers[colIndex];
      headers.splice(colIndex, 1);
      headers.splice(newIndex, 0, headerToMove);

      // Move data in each row
      data.forEach(row => {
         const cellToMove = row[colIndex];
         row.splice(colIndex, 1);
         row.splice(newIndex, 0, cellToMove);
      });

      renderSpreadsheet();

      // Set focus to the moved column
      const newActiveCell = dataBody.querySelectorAll('tr')[activeCell.parentNode.rowIndex - 1].children[newIndex];
      newActiveCell.querySelector('input').focus();
   }

   function addRow(position) {
      if (!activeCell) return;

      const rowIndex = Array.from(activeCell.parentNode.parentNode.children).indexOf(activeCell.parentNode);
      const newRowIndex = position === 'before' ? rowIndex : rowIndex + 1;

      // Add new empty row
      const newRow = Array(headers.length).fill('');
      data.splice(newRowIndex, 0, newRow);

      renderSpreadsheet();

      // Set focus to the new row
      const newActiveCell = dataBody.children[newRowIndex].children[activeCell.cellIndex];
      newActiveCell.querySelector('input').focus();
   }

   function removeRow() {
      if (!activeCell) return;

      const rowIndex = Array.from(activeCell.parentNode.parentNode.children).indexOf(activeCell.parentNode);
      if (data.length <= 1) return; // Don't remove the last row

      data.splice(rowIndex, 1);
      renderSpreadsheet();
   }

   function moveRow(direction) {
      if (!activeCell) return;

      const rowIndex = Array.from(activeCell.parentNode.parentNode.children).indexOf(activeCell.parentNode);
      const newIndex = direction === 'up' ? rowIndex - 1 : rowIndex + 1;

      // Check if move is possible
      if (newIndex < 0 || newIndex >= data.length) return;

      // Move row data
      const rowToMove = data[rowIndex];
      data.splice(rowIndex, 1);
      data.splice(newIndex, 0, rowToMove);

      renderSpreadsheet();

      // Set focus to the moved row
      const newActiveCell = dataBody.children[newIndex].children[activeCell.cellIndex];
      newActiveCell.querySelector('input').focus();
   }

   function sortColumn(colIndex) {
      const currentSort = sortState[colIndex];
      let sortDirection;

      // Determine new sort direction
      if (currentSort === 'asc') {
         sortDirection = 'desc';
      } else if (currentSort === 'desc') {
         sortDirection = null; // No sort
      } else {
         sortDirection = 'asc';
      }

      // Update sort state
      if (sortDirection) {
         sortState = {
            [colIndex]: sortDirection
         };
      } else {
         sortState = {};
      }

      // Sort the data if needed
      if (sortDirection) {
         data.sort((a, b) => {
            const valA = a[colIndex];
            const valB = b[colIndex];

            // Try to compare as numbers
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);

            if (!isNaN(numA) && !isNaN(numB)) {
               return sortDirection === 'asc' ? numA - numB : numB - numA;
            }

            // Fall back to string comparison
            return sortDirection === 'asc' ?
               String(valA).localeCompare(String(valB)) :
               String(valB).localeCompare(String(valA));
         });
      }

      renderSpreadsheet();
   }
});
