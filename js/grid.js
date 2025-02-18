const rows = 150; // No of rows
const cols = 30; // No of columns

// previous Clicked cell
let previousClickedCell = "";

// Cell border Color when click
const cellBorderColor = "#218c74";
const lightGray = "#dfe4ea";

// Collection of whole excel-sheet database
let collectedSheetDB = [];
// Collection of whole sheet-cell  database
let sheetDB = [];

// Copy arrays
let copyData = [];

// Range Storage
let rangeStorage = [];

// Collection of whole excel-sheet
let collecteddGraphComponent = [];
// Collection of whole excel-sheet
let graphComponentMatrix = [];

// Rows and Columns numbering element selection
const addressColCont = document.querySelector(".address-col-cont");
const addressRowCont = document.querySelector(".address-row-cont");
const cellsCont = document.querySelector(".cells-cont");
const addressBar = document.querySelector(".address-bar");

// Rows numbering
for (let i = 0; i < rows; i++) {
  const addressCol = document.createElement("div");
  addressCol.innerText = i + 1;
  addressCol.setAttribute("class", "address-col");
  addressColCont.appendChild(addressCol);
}

// Column numbering
for (let i = 0; i < cols; i++) {
  const addressRow = document.createElement("div");
  if (i < 26) {
    addressRow.innerText = String.fromCharCode(65 + i);
  } else {
    addressRow.innerText = "A" + String.fromCharCode(65 + i - 26);
  }
  addressRow.setAttribute("class", "address-row");
  addressRowCont.appendChild(addressRow);
}

// Create cells
for (let i = 0; i < rows; i++) {
  const rowCont = document.createElement("div");
  rowCont.setAttribute("class", "row-cont");
  for (let j = 0; j < cols; j++) {
    const cell = document.createElement("div");
    cell.setAttribute("class", "cell");
    cell.setAttribute("contenteditable", "true");
    cell.setAttribute("spellcheck", "false");
    cell.setAttribute("rid", i);
    cell.setAttribute("cid", j);
    rowCont.appendChild(cell);
    addClickListenerForCells(cell, i + 1, j + 1);
    addBlurEventListenerOnCell(cell);
  }
  cellsCont.appendChild(rowCont);
}

// Add Listener to cells and update address bar
function addClickListenerForCells(cell, i, j) {
  cell.addEventListener("click", (e) => {
    if (j <= 26) {
      addressBar.value = String.fromCharCode(64 + j) + i;
    } else {
      addressBar.value = "A" + String.fromCharCode(64 + j - 26) + i;
    }
    addListenerToAttachCellProperties(cell);
    handleSelectedCells(cell);
  });
}

// Default cell data
function returnDefauldata() {
  return {
    bold: false,
    italic: false,
    underline: false,
    alignment: "left",
    fontFamily: "Arial",
    fontSize: "14",
    fontColor: "#000000",
    BGcolor: "#ecf0f1",
    value: "",
    formula: "",
    children: [],
  };
}

//create DB and graph for new sheet
function createSheetDB_And_createGraphComponentMatrix() {
  const sheetDB1 = [];
  const graphComponentMatrix1 = [];
  for (let i = 0; i < rows; i++) {
    const sheetRow = [];
    const graphRow = [];
    for (let j = 0; j < cols; j++) {
      sheetRow.push(returnDefauldata());
      graphRow.push([]);
    }
    sheetDB1.push(sheetRow);
    graphComponentMatrix1.push(graphRow);
  }
  collectedSheetDB.push(sheetDB1);
  collecteddGraphComponent.push(graphComponentMatrix1);
}

// Add blur Event listener
let shouldCheck = true;
function addBlurEventListenerOnCell(cell) {
  cell.addEventListener("blur", (e) => {
    const address = addressBar.value;
    const [activedCell, cellprop] = getCellandCellProp(address);
    const enterData = activedCell.innerText;
    if (enterData[0] === "=" && shouldCheck) {
      const inputFormula = enterData.slice(1);
      cellAsFormulaBar(inputFormula);
    } else {
      cellprop.value = enterData;
      removeChildFromParent(cellprop.formula);
      cellprop.formula = "";
      updateChildrenCells(address);
    }
  });
}

async function cellAsFormulaBar(inputFormula) {
  if (inputFormula) {
    const address = addressBar.value;
    const [cell, cellProp] = getCellandCellProp(address);
    if (inputFormula !== cellProp.formula) {
      removeChildFromParent(cellProp.formula);
    }
    const exactFormula = GetExactFormula(inputFormula);
    addChildtoGraphComponent(inputFormula, address);
    const cycleStart = isGraphCyclic(graphComponentMatrix);
    if (cycleStart) {
      let isOk = confirm(
        "There is Cyclic in cells formulas. Do you want to trace your cyclic cell path"
      );
      while (isOk) {
        shouldCheck = false;
        await isGraphCyclicTracePath(graphComponentMatrix, cycleStart);
        isOk = confirm("Do you want to trace your cyclic again");
        cell.innerText = "";
      }
      shouldCheck = true;
      removeChildFromGraphComponent(inputFormula);
      return;
    }
    const evaluatedValue = evaluatedFormula(inputFormula);
    setCellUIAndCellProp(evaluatedValue, exactFormula, address);
    addChildToParent(exactFormula);
    updateChildrenCells(address);
  }
}

// Selection of tool-bar element
const bold = document.querySelector(".bold");
const italic = document.querySelector(".italic");
const underline = document.querySelector(".underline");
const fontSize = document.querySelector(".font-size-prop");
const fontFamily = document.querySelector(".font-family-prop");
const fontColor = document.querySelector(".font-color-prop");
const BGcolor = document.querySelector(".BGcolor-prop");
const alignment = document.querySelectorAll(".alignment");
const leftAlignment = alignment[0];
const centerAlignment = alignment[1];
const rightAlignment = alignment[2];

// Color
const activeColorProp = "#d1d8e0";
const inactiveColorProp = "#ecf0f1";

// Event-Listener to the all tools
{
  bold.addEventListener("click", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    // Change in database
    cellprop.bold = !cellprop.bold;
    //Channge in Cell
    cell.style.fontWeight = cellprop.bold ? "bold" : "normal";
    // Change in tool-bar
    bold.style.backgroundColor = cellprop.bold
      ? activeColorProp
      : inactiveColorProp;
  });

  italic.addEventListener("click", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.italic = !cellprop.italic;
    cell.style.fontStyle = cellprop.italic ? "italic" : "normal";
    italic.style.backgroundColor = cellprop.italic
      ? activeColorProp
      : inactiveColorProp;
  });

  underline.addEventListener("click", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.underline = !cellprop.underline;
    cell.style.textDecoration = cellprop.underline ? "underline" : "none";
    underline.style.backgroundColor = cellprop.underline
      ? activeColorProp
      : inactiveColorProp;
  });

  fontSize.addEventListener("change", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.fontSize = fontSize.value;
    cell.style.fontSize = cellprop.fontSize + "px";
    fontSize.value = cellprop.fontSize;
  });

  fontFamily.addEventListener("change", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.fontFamily = fontFamily.value;
    cell.style.fontFamily = cellprop.fontFamily;
    fontFamily.value = cellprop.fontFamily;
  });

  fontColor.addEventListener("change", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.fontColor = fontColor.value;
    cell.style.color = cellprop.fontColor;
    fontColor.value = cellprop.fontColor;
  });
  BGcolor.addEventListener("change", (e) => {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    cellprop.BGcolor = BGcolor.value;
    cell.style.backgroundColor = cellprop.BGcolor;
    BGcolor.value = cellprop.BGcolor;
  });

  alignment.forEach((alignEle) => {
    alignEle.addEventListener("click", (e) => {
      const address = addressBar.value;
      const [cell, cellprop] = getCellandCellProp(address);

      const alignValue = e.target.classList[0];
      cellprop.alignment = alignValue;
      cell.style.textAlign = cellprop.alignment;

      alignmentWork(cellprop.alignment);
    });
  });
}

// Restore Style of the clicked cell

function addListenerToAttachCellProperties(cell) {
  if (rangeStorage.length === 2) {
    SelectCellborder(1, lightGray);
  }
  const address = addressBar.value;
  const [rid, cid] = decodeRIDCIDfromAddress(address);
  cell.style.border = `3px solid ${cellBorderColor}`;
  const [prid, pcid] = decodeRIDCIDfromAddress(previousClickedCell);
  if (previousClickedCell && (rid !== prid || cid !== pcid)) {
    const Prevcell = document.querySelector(
      `.cell[rid="${prid}"][cid="${pcid}"]`
    );
    Prevcell.style.border = `1px solid #dfe4ea`;
  }
  previousClickedCell = addressBar.value;
  const cellprop = sheetDB[rid][cid];
  cell.style.fontWeight = cellprop.bold ? "bold" : "normal";
  cell.style.fontStyle = cellprop.italic ? "italic" : "normal";
  cell.style.textDecoration = cellprop.underline ? "underline" : "none";
  cell.style.fontSize = cellprop.fontSize + "px";
  cell.style.fontFamily = cellprop.fontFamily;
  cell.style.color = cellprop.fontColor;
  cell.style.backgroundColor = cellprop.BGcolor;
  cell.style.textAlign = cellprop.alignment;

  bold.style.backgroundColor = cellprop.bold
    ? activeColorProp
    : inactiveColorProp;
  italic.style.backgroundColor = cellprop.italic
    ? activeColorProp
    : inactiveColorProp;
  underline.style.backgroundColor = cellprop.underline
    ? activeColorProp
    : inactiveColorProp;
  fontSize.value = cellprop.fontSize;
  fontFamily.value = cellprop.fontFamily;
  fontColor.value = cellprop.fontColor;
  BGcolor.value = cellprop.BGcolor;

  alignmentWork(cellprop.alignment);

  const formulaBar = document.querySelector(".formula-bar");
  formulaBar.value = cellprop.formula;
  cell.innerText = cellprop.value;
}
// Switch UI to the selected alignment icon
function alignmentWork(cellAlignment) {
  switch (cellAlignment) {
    case "left":
      leftAlignment.style.backgroundColor = activeColorProp;
      centerAlignment.style.backgroundColor = inactiveColorProp;
      rightAlignment.style.backgroundColor = inactiveColorProp;
      break;
    case "center":
      leftAlignment.style.backgroundColor = inactiveColorProp;
      centerAlignment.style.backgroundColor = activeColorProp;
      rightAlignment.style.backgroundColor = inactiveColorProp;
      break;
    case "right":
      leftAlignment.style.backgroundColor = inactiveColorProp;
      centerAlignment.style.backgroundColor = inactiveColorProp;
      rightAlignment.style.backgroundColor = activeColorProp;
      break;
  }
}

// Get cell and cell properties
function getCellandCellProp(address) {
  const [rid, cid] = decodeRIDCIDfromAddress(address);
  const cell = document.querySelector(`.cell[rid="${rid}"][cid="${cid}"]`);
  const cellprop = sheetDB[rid][cid];
  return [cell, cellprop];
}

// Decode Rows and Columns
function decodeRIDCIDfromAddress(address) {
  let rid;
  let cid;
  if (
    Number(address.charCodeAt(1)) >= 65 &&
    Number(address.charCodeAt(1)) <= 90
  ) {
    rid = Number(address.slice(2)) - 1;
    cid = Number(address.charCodeAt(1)) - 65 + 26;
  } else {
    rid = Number(address.slice(1)) - 1;
    cid = Number(address.charCodeAt(0)) - 65;
  }
  return [rid, cid];
}

const formulaBar = document.querySelector(".formula-bar");
formulaBar.addEventListener("keydown", async (e) => {
  const inputFormula = formulaBar.value;
  if (e.key === "Enter" && inputFormula) {
    const address = addressBar.value;
    const [, cellProp] = getCellandCellProp(address);
    if (inputFormula !== cellProp.formula) {
      removeChildFromParent(cellProp.formula);
    }
    const exactFormula = GetExactFormula(inputFormula);
    addChildtoGraphComponent(exactFormula, address);
    const cycleStart = isGraphCyclic(graphComponentMatrix);
    if (cycleStart) {
      let isOk = confirm(
        "There is Cyclic in cells formulas. Do you want to trace your cyclic cell path"
      );
      while (isOk) {
        await isGraphCyclicTracePath(graphComponentMatrix, cycleStart);
        isOk = confirm("Do you want to trace your cyclic again");
      }
      removeChildFromGraphComponent(exactFormula);
      return;
    }
    const evaluatedValue = evaluatedFormula(exactFormula);
    setCellUIAndCellProp(evaluatedValue, exactFormula, address);
    addChildToParent(exactFormula);
    updateChildrenCells(address);
  }
});

formulaBar.addEventListener("blur", async (e) => {
  const inputFormula = formulaBar.value
  if (inputFormula) {
    const address = addressBar.value
    const [, cellProp] = getCellandCellProp(address)
    if (inputFormula !== cellProp.formula) {
      removeChildFromParent(cellProp.formula)
    }
    const exactFormula = GetExactFormula(inputFormula);
    addChildtoGraphComponent(exactFormula, address)
    const cycleStart = isGraphCyclic(graphComponentMatrix)
    if (cycleStart) {
      let isOk = confirm("There is Cyclic in cells formulas. Do you want to trace your cyclic cell path")
      while (isOk) {
        await isGraphCyclicTracePath(graphComponentMatrix, cycleStart)
        isOk = confirm("Do you want to trace your cyclic again")
      }
      removeChildFromGraphComponent(exactFormula)
      return;
    }
    const evaluatedValue = evaluatedFormula(exactFormula)
    setCellUIAndCellProp(evaluatedValue, exactFormula, address)
    addChildToParent(exactFormula)
    updateChildrenCells(address)
  }
})

//Remove child cell from parent DB

function removeChildFromParent(formula) {
  const childAddress = addressBar.value;
  const arr = getArrayOfFormula(formula);
  for (let i = 0; i < arr.length; i++) {
    const charect = Number(arr[i].charCodeAt(0));
    if (charect >= 65 && charect <= 90) {
      const [, cellprop] = getCellandCellProp(arr[i]);
      const idx = cellprop.children.indexOf(childAddress);
      cellprop.childAddress.splice(idx, 1);
    }
  }
}

//Update child cell DB and UI
function updateChildrenCells(parentAddress) {
  const [, parentCellProp] = getCellandCellProp(parentAddress);
  const children = parentCellProp.children;
  for (let child of children) {
    const [, childCellProp] = getCellandCellProp(child);
    const childformula = childCellProp.formula;
    const evaluatedValue = evaluatedFormula(childformula);
    setCellUIAndCellProp(evaluatedValue, childformula, child);
    updateChildrenCells(child);
  }
}

// Add child cell to parent DB
function addChildToParent(formula) {
  const childAddress = addressBar.value;
  const arr = getArrayOfFormula(formula);
  for (let i = 0; i < arr.length; i++) {
    const charect = Number(arr[i].charCodeAt(0));
    if (charect >= 65 && charect <= 90) {
      const [, cellprop] = getCellandCellProp(arr[i]);
      cellprop.children.push(childAddress);
    }
  }
}

// Evaluation of formula
function evaluatedFormula(formula) {
  const arr = getArrayOfFormula(formula);
  formula = "";
  for (let i = 0; i < arr.length; i++) {
    const charect = Number(arr[i].charCodeAt(0));
    if (charect >= 65 && charect <= 90) {
      const [, cellprop] = getCellandCellProp(arr[i]);
      const enterData = cellprop.value;
      formula += enterData;
    } else {
      formula += arr[i];
    }
  }
  return eval(formula);
}

// Change in cell DB and UI
function setCellUIAndCellProp(evaluatedValue, formula, address) {
  const [activedCell, cellprop] = getCellandCellProp(address);
  activedCell.innerText = evaluatedValue;
  cellprop.value = evaluatedValue;
  cellprop.formula = formula;
}

// Get array of the formula
function getArrayOfFormula(formula) {
  let str = "";
  const arr = [];
  for (let i = 0; i < formula.length; i++) {
    if (formula[i] == " ") {
      continue;
    } else if (
      formula[i] == "(" ||
      formula[i] == ")" ||
      formula[i] == "{" ||
      formula[i] == "}" ||
      formula[i] == "[" ||
      formula[i] == "]"
    ) {
      arr.push(formula[i]);
    } else if (
      formula[i] == "+" ||
      formula[i] == "-" ||
      formula[i] == "*" ||
      formula[i] == "/"
    ) {
      arr.push(str.toUpperCase());
      arr.push(formula[i]);
      str = "";
    } else {
      str += formula[i];
    }
  }
  arr.push(str.toUpperCase());
  return arr;
}

// To get exact formula
function GetExactFormula(formula) {
  const arr = getArrayOfFormula(formula);
  formula = "";
  for (let i = 0; i < arr.length; i++) {
    formula += arr[i];
  }
  return formula;
}

// Add child in graphDB
function addChildtoGraphComponent(exactFormula, childAddress) {
  const [crid, ccid] = decodeRIDCIDfromAddress(childAddress);
  const arr = getArrayOfFormula(exactFormula);
  for (let i = 0; i < arr.length; i++) {
    const charect = Number(arr[i].charCodeAt(0));
    if (charect >= 65 && charect <= 90) {
      const [prid, pcid] = decodeRIDCIDfromAddress(arr[i]);
      graphComponentMatrix[prid][pcid].push([crid, ccid]);
    }
  }
}

// To find where is graph found;
function isGraphCyclic(graphComponentMatrix) {
  const visited = [];
  const dfsVisited = [];
  for (let i = 0; i < rows; i++) {
    const visitedRow = [];
    const dfsVisitedRow = [];
    for (let j = 0; j < cols; j++) {
      visitedRow.push(false);
      dfsVisitedRow.push(false);
    }
    visited.push(visitedRow);
    dfsVisited.push(dfsVisitedRow);
  }
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      if (!visited[i][j]) {
        const isTrue = dfsCycleDetection(
          graphComponentMatrix,
          i,
          j,
          visited,
          dfsVisited
        );
        if (isTrue) {
          return [i, j];
        }
      }
    }
  }
  return null;
}

// To detect cycle
function dfsCycleDetection(
  graphComponentMatrix,
  srcr,
  srcc,
  visited,
  dfsVisited
) {
  visited[srcr][srcc] = true;
  dfsVisited[srcr][srcc] = true;

  for (let i = 0; i < graphComponentMatrix[srcr][srcc].length; i++) {
    const [crid, ccid] = graphComponentMatrix[srcr][srcc][i];
    if (!visited[crid][ccid]) {
      const isTrue = dfsCycleDetection(
        graphComponentMatrix,
        crid,
        ccid,
        visited,
        dfsVisited
      );
      if (isTrue) {
        return true;
      }
    } else if (dfsVisited[crid][ccid]) {
      return true;
    }
  }
  dfsVisited[srcr][srcc] = false;
  return false;
}

// To remove child in graphDB after cycle detection
function removeChildFromGraphComponent(exactFormula) {
  const arr = getArrayOfFormula(exactFormula);
  for (let i = 0; i < arr.length; i++) {
    let charect = Number(arr[i].charCodeAt(0));
    if (charect >= 65 && charect <= 90) {
      const [prid, pcid] = decodeRIDCIDfromAddress(arr[i]);
      graphComponentMatrix[prid][pcid].pop();
    }
  }
}

// To provide delay function
function colorPromise() {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      resolve();
    }, 1000);
  });
}

// Trace Path
async function isGraphCyclicTracePath(graphComponentMatrix, cycleStart) {
  const [srcr, srcc] = cycleStart;
  const visited = [];
  const dfsVisited = [];
  for (let i = 0; i < rows; i++) {
    const visitedRow = [];
    const dfsVisitedRow = [];
    for (let j = 0; j < cols; j++) {
      visitedRow.push(false);
      dfsVisitedRow.push(false);
    }
    visited.push(visitedRow);
    dfsVisited.push(dfsVisitedRow);
  }
  const isTrue = await dfsCycleDetectionTracePath(
    graphComponentMatrix,
    srcr,
    srcc,
    visited,
    dfsVisited
  );
  if (isTrue) {
    return Promise.resolve(true);
  }
  return Promise.resolve(false);
}
// To detect cycle and color them
async function dfsCycleDetectionTracePath(
  graphComponentMatrix,
  srcr,
  srcc,
  visited,
  dfsVisited
) {
  visited[srcr][srcc] = true;
  dfsVisited[srcr][srcc] = true;
  const cell = document.querySelector(`.cell[rid="${srcr}"][cid="${srcc}"]`);
  cell.style.backgroundColor = "lightblue";
  await colorPromise();
  for (let i = 0; i < graphComponentMatrix[srcr][srcc].length; i++) {
    const [crid, ccid] = graphComponentMatrix[srcr][srcc][i];
    if (!visited[crid][ccid]) {
      const isTrue = await dfsCycleDetectionTracePath(
        graphComponentMatrix,
        crid,
        ccid,
        visited,
        dfsVisited
      );
      if (isTrue) {
        cell.style.backgroundColor = "transparent";
        await colorPromise();
        return Promise.resolve(true);
      }
    } else if (visited[crid][ccid] && dfsVisited[crid][ccid]) {
      const CyclicCell = document.querySelector(
        `.cell[rid="${crid}"][cid="${ccid}"]`
      );
      CyclicCell.style.backgroundColor = "lightsalmon";
      await colorPromise();
      CyclicCell.style.backgroundColor = "transparent";
      await colorPromise();
      cell.style.backgroundColor = "transparent";
      await colorPromise();
      return Promise.resolve(true);
    }
  }
  dfsVisited[srcr][srcc] = false;
  return Promise.resolve(false);
}

const addSheetBtn = document.querySelector(".sheet-add-icon");
const sheetFolderCont = document.querySelector(".sheets-folder-cont");
const activeSheetColor = "#ced6e0";



// click listener to add sheet btn
addSheetBtn.addEventListener("click", (e) => {
  const sheet = document.createElement("div");
  sheet.setAttribute("class", "sheet-folder");
  const allSheetFolder = document.querySelectorAll(".sheet-folder");
  sheet.setAttribute("id", allSheetFolder.length);
  sheet.innerHTML = `
        <div class="sheet-content">Sheet-${allSheetFolder.length + 1}</div>
    `;
  sheetFolderCont.appendChild(sheet);
  createSheetDB_And_createGraphComponentMatrix();
  handleSheetActiveness(sheet);
  handlesheetRemoval(sheet);
  sheet.click();
  sheet.scrollIntoView();
});

function handleSheetActiveness(sheet) {
  sheet.addEventListener("click", (e) => {
    const sheetIndex = Number(sheet.getAttribute("id"));
    handleSheetDB(sheetIndex);
    handleSheetUi(sheet);

  });
}

function handleSheetDB(sheetIndex) {
  sheetDB = collectedSheetDB[sheetIndex];
  graphComponentMatrix = collecteddGraphComponent[sheetIndex];
}

function handleSheetProp() {
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const cell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
      cell.click();
    }
  }
  const firstCell = document.querySelector(".cell");
  firstCell.click();
  firstCell.style.border = `3px solid ${cellBorderColor}`;
}

function handleSheetUi(sheet) {
  const allSheetFolder = document.querySelectorAll(".sheet-folder");
  for (let i = 0; i < allSheetFolder.length; i++) {
    allSheetFolder[i].style.backgroundColor = "transparent";
  }
  sheet.style.backgroundColor = activeSheetColor;
  handleSheetProp();
}

function handlesheetRemoval(sheet) {
  sheet.addEventListener("mousedown", (e) => {
    if (e.button !== 2) {
      return;
    }
    const allSheetFolder = document.querySelectorAll(".sheet-folder");
    if (allSheetFolder.length === 1) {
      alert("You need to have atleast one sheet");
      return;
    }
    const isOk = confirm("Do you want to delete permanently");
    if (!isOk) {
      return;
    }

    const sheetIndex = Number(sheet.getAttribute("id"));
    collectedSheetDB.splice(sheetIndex, 1);
    collecteddGraphComponent.splice(sheetIndex, 1);
    handleSheetUiRemoval(sheet);
    sheetDB = collectedSheetDB[0];
    graphComponentMatrix = collecteddGraphComponent[0];
    handleSheetProp();
  });
}

function handleSheetUiRemoval(sheet) {
  sheet.remove();
  const allSheetFolder = document.querySelectorAll(".sheet-folder");
  for (let i = 0; i < allSheetFolder.length; i++) {
    allSheetFolder[i].setAttribute("id", i);
    const sheetContent = allSheetFolder[i].querySelector(".sheet-content");
    sheetContent.innerText = `Sheet-${i + 1}`;
    allSheetFolder[i].style.backgroundColor = "transparent";
  }
  allSheetFolder[0].style.backgroundColor = activeSheetColor;
}

const copybtn = document.querySelector(".copy");
const pastebtn = document.querySelector(".paste");
const cutbtn = document.querySelector(".cut");

let shiftKey;
document.addEventListener("keydown", (e) => {
  shiftKey = e.shiftKey;
});
document.addEventListener("keyup", (e) => {
  shiftKey = e.shiftKey;
});


// Range selection
function handleSelectedCells(cell) {
  if (!shiftKey) {
    return;
  }
  if (rangeStorage.length >= 2) {
    SelectCellborder(1, lightGray);
    rangeStorage = [];
  }
  const rid = Number(cell.getAttribute("rid"));
  const cid = Number(cell.getAttribute("cid"));
  rangeStorage.push([rid, cid]);
  if (rangeStorage.length == 2) {
    SelectCellborder(3, cellBorderColor);
  }
}
// border color of while selecton and unselection
function SelectCellborder(size, color) {
  for (let i = rangeStorage[0][0]; i <= rangeStorage[1][0]; i++) {
    for (let j = rangeStorage[0][1]; j <= rangeStorage[1][1]; j++) {
      const rangeCell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
      if (i === rangeStorage[0][0] && j === rangeStorage[0][1]) {
        rangeCell.style.borderLeft = `${size}px solid ${color}`;
        rangeCell.style.borderTop = `${size}px solid ${color}`;
      } else if (i === rangeStorage[0][0]) {
        rangeCell.style.borderTop = `${size}px solid ${color}`;
      }
      if (i === rangeStorage[1][0] && j === rangeStorage[0][1]) {
        rangeCell.style.borderLeft = `${size}px solid ${color}`;
        rangeCell.style.borderBottom = `${size}px solid ${color}`;
      } else if (j === rangeStorage[0][1]) {
        rangeCell.style.borderLeft = `${size}px solid ${color}`;
      }
      if (i === rangeStorage[0][0] && j === rangeStorage[1][1]) {
        rangeCell.style.borderRight = `${size}px solid ${color}`;
        rangeCell.style.borderTop = `${size}px solid ${color}`;
      } else if (j === rangeStorage[1][1]) {
        rangeCell.style.borderRight = `${size}px solid ${color}`;
      }
      if (i === rangeStorage[1][0] && j === rangeStorage[1][1]) {
        rangeCell.style.borderRight = `${size}px solid ${color}`;
        rangeCell.style.borderBottom = `${size}px solid ${color}`;
        if (size === 3) {
          rangeCell.style.borderTop = `1px solid ${lightGray}`;
          rangeCell.style.borderLeft = `1px solid ${lightGray}`;
        }
      } else if (i === rangeStorage[1][0]) {
        rangeCell.style.borderBottom = `${size}px solid ${color}`;
      }
    }
  }
}

let singleCellprop = {};
copybtn.addEventListener("click", (e) => {
  copyData = [];
  if (rangeStorage.length < 2) {
    const address = addressBar.value;
    const [, cellprop] = getCellandCellProp(address);
    singleCellprop = cellprop;
    return;
  }

  for (let i = rangeStorage[0][0]; i <= rangeStorage[1][0]; i++) {
    const row = [];
    for (let j = rangeStorage[0][1]; j <= rangeStorage[1][1]; j++) {
      row.push(sheetDB[i][j]);
    }
    copyData.push(row);
  }
  console.log(copyData)
  if (rangeStorage.length === 2) SelectCellborder(1, lightGray);
});

pastebtn.addEventListener("click", (e) => {
  const address = addressBar.value;
  const [rid, cid] = decodeRIDCIDfromAddress(address);
  console.log("ok");
  if (rangeStorage.length < 2) {
    sheetDB[rid][cid] = singleCellprop;
    console.log(sheetDB[rid][cid]);
    const rangeCell = document.querySelector(
      `.cell[rid="${rid}"][cid="${cid}"]`
    );
    rangeCell.click();
    return;
  }
  if (copyData.length === 0) {
    return;
  }
  for (
    let i = rid, k = 0;
    i <= rid + Math.abs(rangeStorage[1][0] - rangeStorage[0][0]);
    k++, i++
  ) {
    for (
      let j = cid, l = 0;
      j <= cid + Math.abs(rangeStorage[1][1] - rangeStorage[0][1]);
      j++, l++
    ) {
      const rangeCell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
      if (rangeCell) {
        sheetDB[i][j] = copyData[k][l];
        rangeCell.click();
      }
    }
  }
});

cutbtn.addEventListener("click", (e) => {
  if (rangeStorage.length < 2) {
    const address = addressBar.value;
    const [cell, cellprop] = getCellandCellProp(address);
    singleCellprop = cellprop;
    cell.innerText = "";
    return;
  }
  copyData = [];
  for (let i = rangeStorage[0][0]; i <= rangeStorage[1][0]; i++) {
    const row = [];
    for (let j = rangeStorage[0][1]; j <= rangeStorage[1][1]; j++) {
      const rangeCell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
      row.push(sheetDB[i][j]);
      sheetDB[i][j] = returnDefauldata();
      rangeCell.click();
    }
    copyData.push(row);
  }
  if (rangeStorage.length === 2) SelectCellborder(1, lightGray);
});

const downloadbtn = document.querySelector(".download");
const uploadbtn = document.querySelector(".upload");

downloadbtn.addEventListener("click", (e) => {
  const data = JSON.stringify([sheetDB, graphComponentMatrix]);
  const file = new Blob([data], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file);
  a.download = "sheetDATA.json";
  a.click();
});

uploadbtn.addEventListener("click", (e) => {
  const input = document.createElement("input");
  input.setAttribute("type", "file");
  input.click();
  input.addEventListener("change", (e) => {
    const fr = new FileReader();
    const files = input.files;
    const fileObj = files[0];
    fr.readAsText(fileObj);
    fr.addEventListener("load", (e) => {
      const readSheetData = JSON.parse(fr.result);
      addSheetBtn.click();
      sheetDB = readSheetData[0];
      graphComponentMatrix = readSheetData[0];
      collectedSheetDB[collectedSheetDB.length - 1] = sheetDB;
      collecteddGraphComponent[collecteddGraphComponent.length - 1] =
        graphComponentMatrix;
      handleSheetProp();
    });
  });
});
{
  const addSheetBtn = document.querySelector(".sheet-add-icon");
  addSheetBtn.click();
}