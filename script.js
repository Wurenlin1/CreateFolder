let folders = [];
let dontAskAgain = false; // Global flag for "Don't ask again" checkbox
let uploadedData = []; // Store uploaded data
let selectedCells = new Set(); // Track selected cells

// Function to reset the file upload modal
function resetFileUploadModal() {
  document.getElementById('fileInput').value = ''; // Clear file input
  document.getElementById('selectedRange').value = ''; // Clear selected range
  document.getElementById('filePreview').innerHTML = ''; // Clear preview grid
  uploadedData = []; // Clear uploaded data
  selectedCells.clear(); // Clear selected cells
  updateSelectedCellCount(); // Reset selected cell count
}

// Function to close the file upload modal
function closeFileUploadModal() {
  const modal = document.getElementById('fileUploadModal');
  modal.style.display = 'none';
  resetFileUploadModal(); // Reset the modal
}

// Function to update the selected cell count
function updateSelectedCellCount() {
  const selectedCellCount = selectedCells.size;
  document.getElementById('selectedCellCount').textContent = selectedCellCount;
}

// Function to check for duplicates in a list of names
function checkForDuplicates(names) {
  const duplicates = [];
  const uniqueNames = new Set();

  names.forEach(name => {
    const trimmedName = name.trim();
    if (uniqueNames.has(trimmedName)) {
      duplicates.push(trimmedName);
    } else {
      uniqueNames.add(trimmedName);
    }
  });

  return duplicates;
}

// Function to generate a unique folder name by appending a counter in parentheses
function generateUniqueFolderName(name, folderList) {
  let newName = name;
  let counter = 1;

  // Check if the name already exists in the folder list
  while (folderList.some(folder => folder.name === newName)) {
    newName = `${name}(${counter})`;
    counter++;
  }

  return newName;
}

// Function to show the custom modal dialog
function showModal(duplicates, parentFolder) {
  const modal = document.getElementById('duplicateModal');
  const duplicateList = document.getElementById('duplicateList');

  // Clear previous content
  duplicateList.innerHTML = '';

  // Add message at the top of the modal
  const message = document.createElement('p');
  message.textContent = 'The following folder names are duplicates. Check or uncheck the boxes to keep the names.';
  duplicateList.appendChild(message);

  // Add checkboxes for each duplicate
  duplicates.forEach(duplicate => {
    const label = document.createElement('label');
    label.innerHTML = `<input type="checkbox" name="duplicate" value="${duplicate}" checked> ${duplicate}`;
    duplicateList.appendChild(label);
  });

  // Show the modal
  modal.style.display = 'flex';

  // Store the parent folder for later use
  modal.dataset.parentFolder = JSON.stringify(parentFolder || null);
}

// Function to close the custom modal dialog
function closeModal() {
  const modal = document.getElementById('duplicateModal');
  modal.style.display = 'none';
}

// Function to handle user selection in the modal
function handleDuplicateSelection() {
  const checkboxes = document.querySelectorAll('#duplicateList input[type="checkbox"]');
  const selectedDuplicates = [];

  checkboxes.forEach(checkbox => {
    if (checkbox.checked) {
      selectedDuplicates.push(checkbox.value);
    }
  });

  // Update the "Don't ask again" flag
  dontAskAgain = document.getElementById('dontAskAgain').checked;

  // Close the modal
  closeModal();

  // Get the parent folder from the modal dataset
  const parentFolder = JSON.parse(document.getElementById('duplicateModal').dataset.parentFolder);

  // Get all folder names from the input
  const folderNames = parentFolder
    ? document.getElementById('subfolderInput').value.trim().split('\n')
    : document.getElementById('mainFolderInput').value.trim().split('\n');

  if (parentFolder) {
    // Process subfolder duplicates
    const selectedFolders = getSelectedFolders(folders);
    selectedFolders.forEach(folder => {
      folderNames.forEach(name => {
        const trimmedName = name.trim();
        if (trimmedName) {
          if (selectedDuplicates.includes(trimmedName)) {
            // Generate a unique name for duplicates
            const uniqueName = generateUniqueFolderName(trimmedName, folder.subfolders);
            folder.subfolders.push({ name: uniqueName, subfolders: [], isSelected: false, isExpanded: true });
          } else if (!folder.subfolders.some(subfolder => subfolder.name === trimmedName)) {
            // Add non-duplicate names
            folder.subfolders.push({ name: trimmedName, subfolders: [], isSelected: false, isExpanded: true });
          }
        }
      });
    });
  } else {
    // Process main folder duplicates
    folderNames.forEach(name => {
      const trimmedName = name.trim();
      if (trimmedName) {
        if (selectedDuplicates.includes(trimmedName)) {
          // Generate a unique name for duplicates
          const uniqueName = generateUniqueFolderName(trimmedName, folders);
          folders.push({ name: uniqueName, subfolders: [], isSelected: false, isExpanded: true });
        } else if (!folders.some(folder => folder.name === trimmedName)) {
          // Add non-duplicate names
          folders.push({ name: trimmedName, subfolders: [], isSelected: false, isExpanded: true });
        }
      }
    });
  }

  // Clear the input
  if (parentFolder) {
    document.getElementById('subfolderInput').value = '';
  } else {
    document.getElementById('mainFolderInput').value = '';
  }

  // Re-render the folder structure and update counts
  renderFolderStructure();
  updateFolderCounts();
}

// Function to process selected duplicates for main folders
function processDuplicates(selectedDuplicates) {
  const mainFolderNames = document.getElementById('mainFolderInput').value.trim().split('\n');

  // Add all valid folder names (both duplicates and non-duplicates)
  mainFolderNames.forEach(name => {
    const trimmedName = name.trim();
    if (trimmedName) {
      // If the name is a selected duplicate, make it unique
      if (selectedDuplicates.includes(trimmedName)) {
        const uniqueName = generateUniqueFolderName(trimmedName, folders);
        folders.push({ name: uniqueName, subfolders: [], isSelected: false, isExpanded: true });
        alert(`Duplicate detected: "${trimmedName}" renamed to "${uniqueName}".`);
      } else if (!folders.some(folder => folder.name === trimmedName)) {
        // If the name is not a duplicate, add it as-is
        folders.push({ name: trimmedName, subfolders: [], isSelected: false, isExpanded: true });
      } else {
        // If the name already exists and is not a selected duplicate, skip it
        alert(`Skipping duplicate: "${trimmedName}" already exists.`);
      }
    }
  });

  document.getElementById('mainFolderInput').value = '';
  renderFolderStructure();
  updateFolderCounts(); // Update counts
}

// Function to process selected subfolder duplicates
function processSubfolderDuplicates(subfolderNames, parentFolder) {
  // Check for duplicates in the input and in the parent folder's subfolders
  const duplicatesInInput = checkForDuplicates(subfolderNames);
  const duplicatesInPreview = subfolderNames.filter(name => 
    parentFolder.subfolders.some(subfolder => subfolder.name === name.trim())
  );

  const allDuplicates = [...new Set([...duplicatesInInput, ...duplicatesInPreview])];

  if (allDuplicates.length > 0 && !dontAskAgain) {
    // Show modal for duplicates
    showModal(allDuplicates, parentFolder);
  } else {
    // If "Don't ask again" is checked, automatically handle duplicates
    subfolderNames.forEach(name => {
      const trimmedName = name.trim();
      if (trimmedName) {
        if (parentFolder.subfolders.some(subfolder => subfolder.name === trimmedName)) {
          // Generate a unique name for duplicates
          const uniqueName = generateUniqueFolderName(trimmedName, parentFolder.subfolders);
          parentFolder.subfolders.push({ name: uniqueName, subfolders: [], isSelected: false, isExpanded: true });
        } else {
          // Add non-duplicate names
          parentFolder.subfolders.push({ name: trimmedName, subfolders: [], isSelected: false, isExpanded: true });
        }
      }
    });

    document.getElementById('subfolderInput').value = '';
    renderFolderStructure();
    updateFolderCounts();
  }
}

// Function to handle duplicates
function handleDuplicates(mainFolderNames) {
  const duplicatesInInput = checkForDuplicates(mainFolderNames);
  const duplicatesInPreview = mainFolderNames.filter(name => folders.some(folder => folder.name === name.trim()));

  const allDuplicates = [...new Set([...duplicatesInInput, ...duplicatesInPreview])];

  if (allDuplicates.length > 0 && !dontAskAgain) {
    showModal(allDuplicates);
  } else {
    // If "Don't ask again" is checked, automatically keep duplicates
    const processedNames = mainFolderNames.map(name => {
      const trimmedName = name.trim();
      if (allDuplicates.includes(trimmedName)) {
        const uniqueName = generateUniqueFolderName(trimmedName, folders);
        alert(`Duplicate detected: "${trimmedName}" renamed to "${uniqueName}".`);
        return uniqueName;
      }
      return trimmedName;
    });

    // Add the processed names to the folders array
    processedNames.forEach(name => {
      if (name.trim()) {
        folders.push({ name: name.trim(), subfolders: [], isSelected: false, isExpanded: true });
      }
    });

    document.getElementById('mainFolderInput').value = '';
    renderFolderStructure();
    updateFolderCounts(); // Update counts
  }
}

// Function to add multiple main folders
function addMainFolders() {
  const mainFolderNames = document.getElementById('mainFolderInput').value.trim().split('\n');
  if (mainFolderNames.length === 0 || mainFolderNames.every(name => name.trim() === '')) {
    alert('Please enter at least one main folder name.');
    return;
  }
  handleDuplicates(mainFolderNames);

  // Automatically sort folders after adding new ones
  const sortOrder = document.getElementById('sortOrder').value;
  if (sortOrder) {
    sortFolderList(folders, sortOrder);
  }

  renderFolderStructure();
  updateFolderCounts();
}

// Function to add subfolders to selected folders
function addSubfolders() {
  const subfolderNames = document.getElementById('subfolderInput').value.trim().split('\n');

  // Check if subfolder input is empty
  if (subfolderNames.length === 0 || subfolderNames.every(name => name.trim() === '')) {
    alert('Please enter at least one subfolder name.');
    return;
  }

  // Get all selected folders
  const selectedFolders = getSelectedFolders(folders);

  // Check if at least one folder is selected
  if (selectedFolders.length === 0) {
    alert('Please select at least one folder to add subfolders.');
    return;
  }

  // Add subfolders to all selected folders
  selectedFolders.forEach(folder => {
    processSubfolderDuplicates(subfolderNames, folder);
  });

  // Automatically sort folders after adding new ones
  const sortOrder = document.getElementById('sortOrder').value;
  if (sortOrder) {
    sortFolderList(folders, sortOrder);
  }

  renderFolderStructure();
  updateFolderCounts();
}

// Recursive function to get all selected folders (including nested subfolders)
function getSelectedFolders(folderList) {
  let selectedFolders = [];
  folderList.forEach(folder => {
    if (folder.isSelected) {
      selectedFolders.push(folder);
    }
    selectedFolders = selectedFolders.concat(getSelectedFolders(folder.subfolders));
  });
  return selectedFolders;
}

// Function to delete selected folders and subfolders
function deleteSelectedFolders() {
  const selectedFolders = getSelectedFolders(folders);
  if (selectedFolders.length === 0) {
    alert('Please select at least one folder to delete.');
    return;
  }
  folders = deleteSelected(folders);
  renderFolderStructure();
  updateFolderCounts(); // Update counts
}

// Recursive function to delete selected folders
function deleteSelected(folderList) {
  return folderList.filter(folder => {
    if (folder.isSelected) {
      return false;
    }
    folder.subfolders = deleteSelected(folder.subfolders);
    return true;
  });
}

// Function to clear all folders
function clearAllFolders() {
  folders = [];
  renderFolderStructure();
  updateFolderCounts(); // Update counts
}

// Function to save folder structure to a file
function saveFolderStructure() {
  const data = JSON.stringify(folders);
  const blob = new Blob([data], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'folder-structure.json';
  a.click();
  URL.revokeObjectURL(url);
}

// Function to load folder structure from a file
function loadFolderStructure() {
  document.getElementById('loadFile').click();
}

function loadFolderStructureFromFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    folders = JSON.parse(e.target.result);
    renderFolderStructure();
    updateFolderCounts(); // Update counts
  };
  reader.readAsText(file);
}

// Function to render the folder structure
function renderFolderStructure() {
  const folderStructureDiv = document.getElementById('folderStructure');
  folderStructureDiv.innerHTML = ''; // Clear previous content

  folders.forEach(folder => {
    renderFolder(folder, folderStructureDiv, 0);
  });
}

// Recursive function to render folders and subfolders
function renderFolder(folder, parentElement, indentLevel) {
  const folderDiv = document.createElement('div');
  folderDiv.className = 'folder-item';
  folderDiv.style.marginLeft = `${indentLevel * 20}px`;

  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.checked = folder.isSelected;
  checkbox.addEventListener('change', () => {
    folder.isSelected = checkbox.checked;
    updateFolderCounts();
  });

  const folderName = document.createElement('span');
  folderName.textContent = folder.name;
  folderName.classList.toggle('expanded', folder.isExpanded);
  folderName.addEventListener('click', () => {
    folder.isExpanded = !folder.isExpanded;
    renderFolderStructure();
  });

  folderDiv.appendChild(checkbox);
  folderDiv.appendChild(folderName);
  parentElement.appendChild(folderDiv);

  // Render subfolders only if the folder is expanded
  if (folder.isExpanded) {
    folder.subfolders.forEach(subfolder => {
      renderFolder(subfolder, parentElement, indentLevel + 1);
    });
  }
}

// Function to update folder counts
function updateFolderCounts() {
  // Total Main Folders in Textbox
  const mainFoldersInTextbox = document.getElementById('mainFolderInput').value.trim().split('\n').filter(line => line.trim()).length;
  document.getElementById('mainFoldersInTextbox').textContent = mainFoldersInTextbox;

  // Total Main Folders
  const totalMainFolders = folders.length;
  document.getElementById('totalMainFolders').textContent = totalMainFolders;

  // Subfolders in Textbox
  const subfoldersInTextbox = document.getElementById('subfolderInput').value.trim().split('\n').filter(line => line.trim()).length;
  document.getElementById('subfoldersInTextbox').textContent = subfoldersInTextbox;

  // Total Subfolders to Create
  const selectedFolders = getSelectedFolders(folders);
  const totalSubfoldersToCreate = selectedFolders.length * subfoldersInTextbox;
  document.getElementById('totalSubfoldersToCreate').textContent = totalSubfoldersToCreate;
}

// Function to select all folders
function selectAllFolders() {
  const allFolders = getAllFolders(folders);
  const isAnyUnselected = allFolders.some(folder => !folder.isSelected);

  // Toggle selection state
  allFolders.forEach(folder => {
    folder.isSelected = isAnyUnselected;
  });

  renderFolderStructure();
  updateFolderCounts();
}

// Recursive function to get all folders (including nested subfolders)
function getAllFolders(folderList) {
  let allFolders = [];
  folderList.forEach(folder => {
    allFolders.push(folder);
    allFolders = allFolders.concat(getAllFolders(folder.subfolders));
  });
  return allFolders;
}

// Function to sort folders
function sortFolders() {
  const sortOrder = document.getElementById('sortOrder').value;
  if (sortOrder) {
    sortFolderList(folders, sortOrder);
    renderFolderStructure();
  }
}

// Recursive function to sort folders and subfolders
function sortFolderList(folderList, order) {
  folderList.sort((a, b) => {
    if (order === 'asc') {
      return a.name.localeCompare(b.name);
    } else {
      return b.name.localeCompare(a.name);
    }
  });

  // Sort subfolders recursively
  folderList.forEach(folder => {
    if (folder.subfolders.length > 0) {
      sortFolderList(folder.subfolders, order);
    }
  });
}

// Function to toggle expand/collapse all folders
function toggleExpandCollapse() {
  const action = document.getElementById('expandCollapse').value;
  if (action === 'expand' || action === 'collapse') {
    const expand = action === 'expand';
    toggleAllFolders(folders, expand);
    renderFolderStructure();
  }
}

// Recursive function to toggle all folders
function toggleAllFolders(folderList, expand) {
  folderList.forEach(folder => {
    folder.isExpanded = expand;
    if (folder.subfolders.length > 0) {
      toggleAllFolders(folder.subfolders, expand);
    }
  });
}

// Function to select folders by level
function selectFoldersByLevel() {
  const level = parseInt(document.getElementById('folderLevel').value);
  if (isNaN(level)) {
    alert('Please enter a valid folder level.');
    return;
  }

  const foldersAtLevel = getFoldersAtLevel(folders, level, 1);
  if (foldersAtLevel.length === 0) {
    alert(`No folders exist at level ${level}.`);
    return;
  }

  selectFoldersAtLevel(folders, level, 1);
  renderFolderStructure();
  updateFolderCounts();
}

// Recursive function to get folders at a specific level
function getFoldersAtLevel(folderList, targetLevel, currentLevel) {
  let foldersAtLevel = [];
  folderList.forEach(folder => {
    if (currentLevel === targetLevel) {
      foldersAtLevel.push(folder);
    } else {
      foldersAtLevel = foldersAtLevel.concat(getFoldersAtLevel(folder.subfolders, targetLevel, currentLevel + 1));
    }
  });
  return foldersAtLevel;
}

// Recursive function to select folders at a specific level
function selectFoldersAtLevel(folderList, targetLevel, currentLevel) {
  folderList.forEach(folder => {
    if (currentLevel === targetLevel) {
      folder.isSelected = true;
    } else {
      selectFoldersAtLevel(folder.subfolders, targetLevel, currentLevel + 1);
    }
  });
}

// Function to export folders as a ZIP file
function exportFoldersAsZip() {
  const zip = new JSZip();

  // Recursive function to add folders and subfolders to the ZIP
  function addFoldersToZip(zip, folderList, path = '') {
    folderList.forEach(folder => {
      const folderPath = path ? `${path}/${folder.name}` : folder.name;
      zip.folder(folderPath);
      if (folder.subfolders.length > 0) {
        addFoldersToZip(zip, folder.subfolders, folderPath);
      }
    });
  }

  // Add folders to the ZIP
  addFoldersToZip(zip, folders);

  // Generate the ZIP file and trigger download
  zip.generateAsync({ type: 'blob' })
    .then(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'folder-structure.zip';
      a.click();
      URL.revokeObjectURL(url);
    });
}

// File Upload Functions
function openFileUploadModal() {
  const modal = document.getElementById('fileUploadModal');
  modal.style.display = 'flex';
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = e.target.result;
    if (file.name.endsWith('.csv')) {
      parseCSV(data);
    } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      parseExcel(data);
    }
  };
  reader.readAsBinaryString(file);
}

function parseCSV(data) {
  Papa.parse(data, {
    header: true,
    dynamicTyping: true,
    skipEmptyLines: true,
    complete: function (results) {
      uploadedData = results.data;
      renderPreviewGrid(results.meta.fields, results.data);
    },
  });
}

function parseExcel(data) {
  const workbook = XLSX.read(data, { type: 'binary' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const headers = jsonData[0];
  const rows = jsonData.slice(1);
  uploadedData = rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });

  renderPreviewGrid(headers, uploadedData);
}

// Function to render the preview grid with column labels, row labels, and Select All button
function renderPreviewGrid(headers, data) {
  const previewDiv = document.getElementById('filePreview');
  let html = '<table><thead>';

  // First row: Column labels (A, B, C, etc.)
  html += '<tr>';
  html += `<th><button onclick="toggleSelectAll()">Select All</button></th>`; // Select All button
  for (let col = 0; col < headers.length; col++) {
    const columnLabel = indexToColumnLabel(col);
    html += `<th>${columnLabel}</th>`; // Column labels (A, B, C, etc.)
  }
  html += '</tr>';

  // Second row: Actual headers (Name, Email, etc.)
  html += '<tr>';
  html += `<th></th>`; // Empty cell for the Select All column
  headers.forEach(header => {
    html += `<th>${header}</th>`; // Actual headers (Name, Email, etc.)
  });
  html += '</tr>';

  html += '</thead><tbody>';

  // Add row labels (1, 2, 3, ...) and data
  data.forEach((row, rowIndex) => {
    html += `<tr>`;
    // Add row label
    html += `<td class="row-label">${rowIndex + 1}</td>`;

    // Add data cells
    headers.forEach((header, colIndex) => {
      const cellId = `${rowIndex}-${colIndex}`;
      html += `<td id="${cellId}" onclick="toggleCellSelection('${cellId}')">${row[header]}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  previewDiv.innerHTML = html;
}

// Helper function to convert column index to label (0 -> A, 1 -> B, ..., 26 -> AA, etc.)
function indexToColumnLabel(index) {
  let label = '';
  while (index >= 0) {
    label = String.fromCharCode(65 + (index % 26)) + label;
    index = Math.floor(index / 26) - 1;
  }
  return label;
}

// Function to toggle Select All
function toggleSelectAll() {
  const allCells = document.querySelectorAll('.file-preview td:not(.row-label)');
  const isAnyUnselected = Array.from(allCells).some(cell => !cell.classList.contains('selected'));

  allCells.forEach(cell => {
    if (isAnyUnselected) {
      cell.classList.add('selected');
      selectedCells.add(cell.id);
    } else {
      cell.classList.remove('selected');
      selectedCells.delete(cell.id);
    }
  });

  updateSelectedCellCount();
  updateSelectedRangeTextbox();
}

// Function to toggle cell selection
function toggleCellSelection(cellId) {
  const cell = document.getElementById(cellId);
  if (selectedCells.has(cellId)) {
    selectedCells.delete(cellId);
    cell.classList.remove('selected');
  } else {
    selectedCells.add(cellId);
    cell.classList.add('selected');
  }

  // Update the selected range textbox
  updateSelectedRangeTextbox();

  // Update the selected cell count
  updateSelectedCellCount();
}

// Function to update the selected range textbox
function updateSelectedRangeTextbox() {
  if (selectedCells.size === 0) {
    document.getElementById('selectedRange').value = '';
    return;
  }

  // Convert selected cells to their cell references (e.g., A1, B4)
  const cellReferences = Array.from(selectedCells).map(cellId => {
    const [rowIndex, colIndex] = cellId.split('-');
    const columnLetter = String.fromCharCode(65 + parseInt(colIndex)); // Convert column index to letter (A, B, C, etc.)
    const rowNumber = parseInt(rowIndex) + 1; // Convert row index to row number (1-based)
    return `${columnLetter}${rowNumber}`;
  });

  // Group consecutive cells into ranges
  const groupedRanges = groupCellsIntoRanges(cellReferences);

  // Update the selected range textbox with comma-separated cell references and ranges
  document.getElementById('selectedRange').value = groupedRanges.join(',');
}

// Helper function to group consecutive cells into ranges
function groupCellsIntoRanges(cellReferences) {
  if (cellReferences.length === 0) return [];

  // Sort cell references by column and row
  cellReferences.sort((a, b) => {
    const colA = a.match(/[A-Za-z]+/)[0];
    const rowA = parseInt(a.match(/\d+/)[0]);
    const colB = b.match(/[A-Za-z]+/)[0];
    const rowB = parseInt(b.match(/\d+/)[0]);

    if (colA === colB) {
      return rowA - rowB;
    } else {
      return colA.localeCompare(colB);
    }
  });

  const ranges = [];
  let startCell = cellReferences[0];
  let prevCell = startCell;

  for (let i = 1; i < cellReferences.length; i++) {
    const currentCell = cellReferences[i];
    const prevCol = prevCell.match(/[A-Za-z]+/)[0];
    const prevRow = parseInt(prevCell.match(/\d+/)[0]);
    const currentCol = currentCell.match(/[A-Za-z]+/)[0];
    const currentRow = parseInt(currentCell.match(/\d+/)[0]);

    // Check if the current cell is consecutive with the previous cell
    if (currentCol === prevCol && currentRow === prevRow + 1) {
      prevCell = currentCell;
    } else {
      // If not consecutive, add the range or individual cell
      if (startCell === prevCell) {
        ranges.push(startCell);
      } else {
        ranges.push(`${startCell}:${prevCell}`);
      }
      startCell = currentCell;
      prevCell = currentCell;
    }
  }

  // Add the last range or individual cell
  if (startCell === prevCell) {
    ranges.push(startCell);
  } else {
    ranges.push(`${startCell}:${prevCell}`);
  }

  return ranges;
}

// Function to highlight cells from the selected range
function highlightCellsFromRange() {
  const rangeInput = document.getElementById('selectedRange').value.trim();
  if (!rangeInput) return;

  // Clear previous selections
  selectedCells.clear();
  document.querySelectorAll('.file-preview td').forEach(cell => {
    cell.classList.remove('selected');
  });

  // Split the input by commas to handle individual cells and ranges
  const ranges = rangeInput.split(',').map(ref => ref.trim());

  // Highlight cells without validation
  ranges.forEach(ref => {
    if (ref.includes(':')) {
      // Handle range (e.g., A1:E1)
      const [startCell, endCell] = ref.split(':');
      if (!startCell || !endCell) return;

      const startCol = startCell.match(/[A-Za-z]+/)[0].toUpperCase();
      const startRow = parseInt(startCell.match(/\d+/)[0]) || 1;
      const endCol = endCell.match(/[A-Za-z]+/)[0].toUpperCase();
      const endRow = parseInt(endCell.match(/\d+/)[0]) || uploadedData.length;

      const headers = Object.keys(uploadedData[0]);

      for (let row = startRow - 1; row < endRow; row++) {
        for (let col = columnToIndex(startCol); col <= columnToIndex(endCol); col++) {
          if (uploadedData[row] && headers[col]) {
            const cellId = `${row}-${col}`;
            const cell = document.getElementById(cellId);
            if (cell) {
              cell.classList.add('selected');
              selectedCells.add(cellId); // Add to the set (automatically handles duplicates)
            }
          }
        }
      }
    } else {
      // Handle individual cell (e.g., A1)
      const col = ref.match(/[A-Za-z]+/)?.[0].toUpperCase();
      const row = ref.match(/\d+/)?.[0];
      if (!col || !row) return;

      const colIndex = columnToIndex(col);
      const rowIndex = parseInt(row) - 1;

      if (uploadedData[rowIndex] && uploadedData[rowIndex][Object.keys(uploadedData[rowIndex])[colIndex]]) {
        const cellId = `${rowIndex}-${colIndex}`;
        const cell = document.getElementById(cellId);
        if (cell) {
          cell.classList.add('selected');
          selectedCells.add(cellId); // Add to the set (automatically handles duplicates)
        }
      }
    }
  });

  // Update the selected cell count
  updateSelectedCellCount();
}

// Function to load selected cells
function loadSelectedCells() {
  const rangeInput = document.getElementById('selectedRange').value.trim();
  let selectedData = [];

  if (rangeInput) {
    // Split the input by commas to handle individual cells and ranges
    const ranges = rangeInput.split(',').map(ref => ref.trim());

    // Check if any range is invalid
    for (const ref of ranges) {
      if (ref.includes(':')) {
        if (!isRangeValid(ref)) {
          alert(`The range "${ref}" is invalid. Please check your input.`);
          return;
        }
      } else {
        // Handle individual cell (e.g., A1)
        const col = ref.match(/[A-Za-z]+/)?.[0].toUpperCase();
        const row = ref.match(/\d+/)?.[0];
        if (!col || !row) {
          alert(`The cell "${ref}" is invalid. Please check your input.`);
          return;
        }

        const colIndex = columnToIndex(col);
        const rowIndex = parseInt(row) - 1;

        if (
          rowIndex < 0 || rowIndex >= uploadedData.length ||
          colIndex < 0 || colIndex >= Object.keys(uploadedData[0] || {}).length
        ) {
          alert(`The cell "${ref}" is invalid. Please check your input.`);
          return;
        }
      }
    }

    // Process valid ranges and cells
    ranges.forEach(ref => {
      if (ref.includes(':')) {
        // Handle range (e.g., A1:E1)
        const [startCell, endCell] = ref.split(':');
        const startCol = startCell.match(/[A-Za-z]+/)[0].toUpperCase();
        const startRow = parseInt(startCell.match(/\d+/)[0]) || 1;
        const endCol = endCell.match(/[A-Za-z]+/)[0].toUpperCase();
        const endRow = parseInt(endCell.match(/\d+/)[0]) || uploadedData.length;

        const headers = Object.keys(uploadedData[0]);

        for (let row = startRow - 1; row < endRow; row++) {
          for (let col = columnToIndex(startCol); col <= columnToIndex(endCol); col++) {
            if (uploadedData[row] && headers[col]) {
              const cellValue = uploadedData[row][headers[col]];
              if (!selectedData.includes(cellValue)) {
                selectedData.push(cellValue); // Add unique values only
              }
            }
          }
        }
      } else {
        // Handle individual cell (e.g., A1)
        const col = ref.match(/[A-Za-z]+/)?.[0].toUpperCase();
        const row = ref.match(/\d+/)?.[0];
        const colIndex = columnToIndex(col);
        const rowIndex = parseInt(row) - 1;

        if (uploadedData[rowIndex] && uploadedData[rowIndex][Object.keys(uploadedData[rowIndex])[colIndex]]) {
          const cellValue = uploadedData[rowIndex][Object.keys(uploadedData[rowIndex])[colIndex]];
          if (!selectedData.includes(cellValue)) {
            selectedData.push(cellValue); // Add unique values only
          }
        }
      }
    });
  } else {
    // Load manually selected cells
    selectedCells.forEach(cellId => {
      const [rowIndex, colIndex] = cellId.split('-');
      const row = uploadedData[rowIndex];
      const headers = Object.keys(row);
      const cellValue = row[headers[colIndex]];
      if (!selectedData.includes(cellValue)) {
        selectedData.push(cellValue); // Add unique values only
      }
    });
  }

  if (selectedData.length === 0) {
    alert('No cells selected. Please select cells or enter a valid range.');
    return;
  }

  const textarea = document.getElementById('mainFolderInput');
  textarea.value = selectedData.join('\n');
  closeFileUploadModal();
  updateFolderCounts();
}

// Function to check if a range is valid
function isRangeValid(range) {
  const [startCell, endCell] = range.split(':');
  if (!startCell || !endCell) return false;

  const startCol = startCell.match(/[A-Za-z]+/)[0].toUpperCase();
  const startRow = parseInt(startCell.match(/\d+/)[0]);
  const endCol = endCell.match(/[A-Za-z]+/)[0].toUpperCase();
  const endRow = parseInt(endCell.match(/\d+/)[0]);

  // Check if the range is within the bounds of the uploaded data
  const headers = Object.keys(uploadedData[0] || {});
  const maxRow = uploadedData.length;
  const maxCol = headers.length;

  const startColIndex = columnToIndex(startCol);
  const endColIndex = columnToIndex(endCol);

  if (
    startRow < 1 || startRow > maxRow ||
    endRow < 1 || endRow > maxRow ||
    startColIndex < 0 || startColIndex >= maxCol ||
    endColIndex < 0 || endColIndex >= maxCol
  ) {
    return false;
  }

  return true;
}

// Helper function to convert column letters to index (e.g., A -> 0, B -> 1)
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Add event listener to the selected range textbox
document.getElementById('selectedRange').addEventListener('input', highlightCellsFromRange);

// Call updateFolderCounts whenever the main folder input changes
document.getElementById('mainFolderInput').addEventListener('input', updateFolderCounts);

// Call updateFolderCounts whenever the subfolder input changes
document.getElementById('subfolderInput').addEventListener('input', updateFolderCounts);

// Initial render and counts update
renderFolderStructure();
updateFolderCounts();
