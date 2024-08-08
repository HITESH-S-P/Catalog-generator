$("#gs").click(() => tryCatch(gs));
$("#cf").click(() => tryCatch(cf));
$("#mi").click(() => tryCatch(mi));
$("#ft").click(() => tryCatch(ft));
const headingDropdown = document.getElementById(
  "headingDropdown"
) as HTMLSelectElement;
const subHeadingDropdown = document.getElementById(
  "subHeadingDropdown"
) as HTMLSelectElement;
const subSubHeading = document.getElementById(
  "subSubHeading"
) as HTMLSelectElement;
let count = 0;
let subCount;
let subSubCount;
///////////////////////GENERAL-SPECIFICATION/////////////////////////
async function gs() {
  await Word.run(async (context) => {
    const body = context.document.body;

    const wrdTbl = body.insertTable(14, 4, Word.RangeLocation.end);

    const wrdRows = wrdTbl.rows;
    const row_1 = wrdRows.getFirst();
    row_1.shadingColor = "#D9D9D9";
    row_1.font.name = "Arial Bold";
    row_1.font.size = 14;
    row_1.horizontalAlignment = "Centered";
    row_1.verticalAlignment = "Center";
    row_1.values = [
      [
        "General Information Items",
        "Specification\n Main Panel",
        "Unit",
        "Note",
      ],
    ];
    wrdRows.load("items");
    await context.sync();
    wrdRows.items[1].values = [["Display area (AA)", " ", "mm", "-"]];
    wrdRows.items[1].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].font.name = "Times New Roman";
    wrdRows.items[1].preferredHeight = 20;
    wrdRows.items[1].font.size = 12;
    wrdRows.items[1].cells.load("items");
    await context.sync();
    wrdRows.items[1].cells.items[0].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[3].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[0].columnWidth = 196;
    wrdRows.items[1].cells.items[1].columnWidth = 304;
    wrdRows.items[2].values = [["Driver element", " ", "-", "-"]];
    wrdRows.items[2].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].font.name = "Times New Roman";
    wrdRows.items[2].preferredHeight = 20;
    wrdRows.items[2].font.size = 12;
    wrdRows.items[2].cells.load("items");
    await context.sync();
    wrdRows.items[2].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[2].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[3].values = [["Display colors", " ", "colors", "-"]];
    wrdRows.items[3].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[3].horizontalAlignment = "Centered";
    wrdRows.items[3].font.name = "Times New Roman";
    wrdRows.items[3].preferredHeight = 20;
    wrdRows.items[3].font.size = 12;
    wrdRows.items[3].cells.load("items");
    await context.sync();
    wrdRows.items[3].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[3].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[4].values = [["Number of pixels", " ", "dots", "-"]];
    wrdRows.items[4].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[4].horizontalAlignment = "Centered";
    wrdRows.items[4].font.name = "Times New Roman";
    wrdRows.items[4].preferredHeight = 20;
    wrdRows.items[4].font.size = 12;
    wrdRows.items[4].cells.load("items");
    await context.sync();
    wrdRows.items[4].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[4].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[5].values = [["Pixel pitch", " ", "-", "-"]];
    wrdRows.items[5].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[5].horizontalAlignment = "Centered";
    wrdRows.items[5].font.name = "Times New Roman";
    wrdRows.items[5].preferredHeight = 20;
    wrdRows.items[5].font.size = 12;
    wrdRows.items[5].cells.load("items");
    await context.sync();
    wrdRows.items[5].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[5].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[5].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[5].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[6].values = [["Viewing angle", " ", "mm", "-"]];
    wrdRows.items[6].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[6].horizontalAlignment = "Centered";
    wrdRows.items[6].font.name = "Times New Roman";
    wrdRows.items[6].preferredHeight = 20;
    wrdRows.items[6].font.size = 12;
    wrdRows.items[6].cells.load("items");
    await context.sync();
    wrdRows.items[6].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[6].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[6].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[1].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10);
    wrdRows.items[7].values = [["Controller IC", " ", "o'clock", "-"]];
    wrdRows.items[7].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[7].horizontalAlignment = "Centered";
    wrdRows.items[7].font.name = "Times New Roman";
    wrdRows.items[7].preferredHeight = 20;
    wrdRows.items[7].font.size = 12;
    wrdRows.items[7].cells.load("items");
    await context.sync();
    wrdRows.items[7].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[7].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[7].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[1].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[8].values = [["LCM Interface", " ", "-", "-"]];
    wrdRows.items[8].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[8].horizontalAlignment = "Centered";
    wrdRows.items[8].font.name = "Times New Roman";
    wrdRows.items[8].preferredHeight = 20;
    wrdRows.items[8].font.size = 12;
    wrdRows.items[8].cells.load("items");
    await context.sync();
    wrdRows.items[8].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[8].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[8].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[1].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[9].values = [["Operating temperature", " ", "-", "-"]];
    wrdRows.items[9].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[9].horizontalAlignment = "Centered";
    wrdRows.items[9].font.name = "Times New Roman";
    wrdRows.items[9].preferredHeight = 20;
    wrdRows.items[9].font.size = 12;
    wrdRows.items[9].cells.load("items");
    await context.sync();
    wrdRows.items[9].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[9].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[9].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[9].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[10].values = [["Display mode", " ", "-", "-"]];
    wrdRows.items[10].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[10].horizontalAlignment = "Centered";
    wrdRows.items[10].font.name = "Times New Roman";
    wrdRows.items[10].preferredHeight = 20;
    wrdRows.items[10].font.size = 12;
    wrdRows.items[10].cells.load("items");
    await context.sync();
    wrdRows.items[10].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[10].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[10].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[10].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[11].values = [["Operating temperature", " ", "℃ ", "-"]];
    wrdRows.items[11].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[11].horizontalAlignment = "Centered";
    wrdRows.items[11].font.name = "Times New Roman";
    wrdRows.items[11].preferredHeight = 20;
    wrdRows.items[11].font.size = 12;
    wrdRows.items[11].cells.load("items");
    await context.sync();
    wrdRows.items[11].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[11].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[11].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[11].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[12].values = [["Storage temperature", " ", "℃ ", "-"]];
    wrdRows.items[12].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[12].horizontalAlignment = "Centered";
    wrdRows.items[12].font.name = "Times New Roman";
    wrdRows.items[12].preferredHeight = 20;
    wrdRows.items[12].font.size = 12;
    wrdRows.items[12].cells.load("items");
    await context.sync();
    wrdRows.items[12].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[12].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[12].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[12].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[13].values = [["Module bonding technology", " ", "-", "-"]];
    wrdRows.items[13].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[13].horizontalAlignment = "Centered";
    wrdRows.items[13].font.name = "Times New Roman";

    wrdRows.items[13].preferredHeight = 20;
    wrdRows.items[13].font.size = 12;
    wrdRows.items[13].cells.load("items");
    await context.sync();
    wrdRows.items[13].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[13].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[13].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[13].cells.items[1].setCellPadding(Word.CellPaddingLocation.left, 250)
    const tblBorders = wrdTbl.getBorder(Word.BorderLocation.all);

    tblBorders.type = Word.BorderType.single;
    const tblBordersIV = wrdTbl.getBorder(Word.BorderLocation.insideVertical);
    tblBordersIV.type = "Single";
    const tblBordersOT = wrdTbl.getBorder(Word.BorderLocation.outside);
    tblBordersOT.type = "Single";
    wrdTbl.load("headerRowCount");
    wrdTbl.styleLastColumn = true;
    const specificationCell = wrdTbl.getCell(0, 1);
    specificationCell.split(1, 2);
    await context.sync();
  });
}
////////////////////////////////////////////////////////////////////////

//////////////////////////////CTP-FEATURES//////////////////////////////
async function cf() {
  await Word.run(async (context) => {
    const body = context.document.body;

    const wrdTbl = body.insertTable(8, 4, Word.RangeLocation.end);

    const wrdRows = wrdTbl.rows;
    const row_1 = wrdRows.getFirst();
    row_1.shadingColor = "#D9D9D9";
    row_1.font.name = "Arial Bold";
    row_1.font.size = 14;
    row_1.horizontalAlignment = "Centered";
    row_1.verticalAlignment = "Center";
    row_1.values = [
      [
        "General Information Items",
        "Specification\n Main Panel",
        "Unit",
        "Note",
      ],
    ];
    wrdRows.load("items");
    await context.sync();
    wrdRows.items[1].values = [["Resolution", " ", "", "-"]];
    wrdRows.items[1].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].font.name = "Times New Roman";
    wrdRows.items[1].preferredHeight = 20;
    wrdRows.items[1].font.size = 12;
    wrdRows.items[1].cells.load("items");
    await context.sync();
    wrdRows.items[1].cells.items[0].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[3].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[0].columnWidth = 196;
    wrdRows.items[1].cells.items[1].columnWidth = 304;
    wrdRows.items[2].values = [["Structure", " ", "", "-"]];
    wrdRows.items[2].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].font.name = "Times New Roman";
    wrdRows.items[2].preferredHeight = 20;
    wrdRows.items[2].font.size = 12;
    wrdRows.items[2].cells.load("items");
    await context.sync();
    wrdRows.items[2].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[2].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[3].values = [["Controller IC", " ", "", "-"]];
    wrdRows.items[3].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[3].horizontalAlignment = "Centered";
    wrdRows.items[3].font.name = "Times New Roman";
    wrdRows.items[3].preferredHeight = 20;
    wrdRows.items[3].font.size = 12;
    wrdRows.items[3].cells.load("items");
    await context.sync();
    wrdRows.items[3].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[3].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[4].values = [["Interface", " ", "", "-"]];
    wrdRows.items[4].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[4].horizontalAlignment = "Centered";
    wrdRows.items[4].font.name = "Times New Roman";
    wrdRows.items[4].preferredHeight = 20;
    wrdRows.items[4].font.size = 12;
    wrdRows.items[4].cells.load("items");
    await context.sync();
    wrdRows.items[4].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[4].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[5].values = [["Slave Address", " ", "", "-"]];
    wrdRows.items[5].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[5].horizontalAlignment = "Centered";
    wrdRows.items[5].font.name = "Times New Roman";
    wrdRows.items[5].preferredHeight = 20;
    wrdRows.items[5].font.size = 12;
    wrdRows.items[5].cells.load("items");
    await context.sync();
    wrdRows.items[5].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[5].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[5].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[5].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[6].values = [["Touch mode", " ", "", "-"]];
    wrdRows.items[6].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[6].horizontalAlignment = "Centered";
    wrdRows.items[6].font.name = "Times New Roman";
    wrdRows.items[6].preferredHeight = 20;
    wrdRows.items[6].font.size = 12;
    wrdRows.items[6].cells.load("items");
    await context.sync();
    wrdRows.items[6].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[6].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[6].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[1].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10);
    wrdRows.items[7].values = [["Logic level", " ", "", "-"]];
    wrdRows.items[7].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[7].horizontalAlignment = "Centered";
    wrdRows.items[7].font.name = "Times New Roman";
    wrdRows.items[7].preferredHeight = 20;
    wrdRows.items[7].font.size = 12;
    wrdRows.items[7].cells.load("items");
    await context.sync();
    wrdRows.items[7].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[7].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[7].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[1].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    const tblBorders = wrdTbl.getBorder(Word.BorderLocation.all);

    tblBorders.type = Word.BorderType.single;
    const tblBordersIV = wrdTbl.getBorder(Word.BorderLocation.insideVertical);
    tblBordersIV.type = "Single";
    const tblBordersOT = wrdTbl.getBorder(Word.BorderLocation.outside);
    tblBordersOT.type = "Single";
    wrdTbl.load("headerRowCount");
    wrdTbl.styleLastColumn = true;
    const specificationCell = wrdTbl.getCell(0, 1);
    specificationCell.split(1, 2);
    await context.sync();
  });
}

////////////////////////////////////////////////////////////////////////

//////////////////////Mechanical Information///////////////////////////
async function mi() {
  await Word.run(async (context) => {
    const body = context.document.body;

    const wrdTbl = body.insertTable(5, 7, Word.RangeLocation.end);

    const wrdRows = wrdTbl.rows;
    const row_1 = wrdRows.getFirst();
    row_1.shadingColor = "#D9D9D9";
    row_1.font.name = "Arial Bold";
    row_1.font.size = 14;
    row_1.horizontalAlignment = "Centered";
    row_1.verticalAlignment = "Center";
    row_1.values = [["Item", "", "Min.", "Typ.", "Max.", "Unit", "Note"]];
    wrdRows.load("items");
    await context.sync();
    wrdRows.items[1].values = [["", "Horizontal(H)", "", "", "", "mm", "-"]];
    wrdRows.items[1].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].font.name = "Times New Roman";
    wrdRows.items[1].preferredHeight = 20;
    wrdRows.items[1].font.size = 12;
    wrdRows.items[1].cells.load("items");
    await context.sync();
    wrdRows.items[1].cells.items[0].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[3].horizontalAlignment = "Centered";
    wrdRows.items[1].cells.items[0].columnWidth = 241;
    wrdRows.items[1].cells.items[1].columnWidth = 90;
    wrdRows.items[2].values = [["", "Vertical(V)", "", "", "", "mm", "-"]];
    wrdRows.items[2].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].font.name = "Times New Roman";
    wrdRows.items[2].preferredHeight = 20;
    wrdRows.items[2].font.size = 12;
    wrdRows.items[2].cells.load("items");
    await context.sync();
    wrdRows.items[2].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[2].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[2].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[3].values = [
      ["Module size", "Depth(D)", "", "", "", "mm", "-"],
    ];
    wrdRows.items[3].verticalAlignment = Word.VerticalAlignment.center;
    wrdRows.items[3].horizontalAlignment = "Centered";
    wrdRows.items[3].font.name = "Times New Roman";
    wrdRows.items[3].preferredHeight = 20;
    wrdRows.items[3].font.size = 12;
    wrdRows.items[3].cells.load("items");
    await context.sync();
    wrdRows.items[3].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[3].cells.items[3].horizontalAlignment = "Centered";
    wrdTbl.mergeCells(1, 0, 3, 0);
    wrdTbl.mergeCells(0, 0, 0, 1);
    //wrdRows.items[3].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    wrdRows.items[4].values = [["Weight", "", " ", "", "", "g", "-"]];
    wrdRows.items[4].verticalAlignment = "Center";
    wrdRows.items[4].horizontalAlignment = "Centered";
    wrdRows.items[4].font.name = "Times New Roman";
    wrdRows.items[4].preferredHeight = 20;
    wrdRows.items[4].font.size = 12;
    wrdRows.items[4].cells.load("items");
    await context.sync();
    wrdRows.items[4].cells.items[1].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[2].horizontalAlignment = "Centered";
    wrdRows.items[4].cells.items[3].horizontalAlignment = "Centered";
    //wrdRows.items[4].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    const tblBorders = wrdTbl.getBorder(Word.BorderLocation.all);

    tblBorders.type = Word.BorderType.single;
    const tblBordersIV = wrdTbl.getBorder(Word.BorderLocation.insideVertical);
    tblBordersIV.type = "Single";
    const tblBordersOT = wrdTbl.getBorder(Word.BorderLocation.outside);
    tblBordersOT.type = "Single";
    wrdTbl.load("headerRowCount");
    wrdTbl.styleLastColumn = true;
    wrdTbl.mergeCells(4, 0, 4, 1);
    await context.sync();
  });
}

////////////////////////////////////////////////////////////////////////

////////////////////////////////TABLE-FREE/////////////////////////////

async function ft() {
  await Word.run(async (context) => {
    // Prompt the user for the number of rows and columns
    const numRows = prompt("ENTER THE NUMBER OF ROWS:");
    const numColumns = prompt("ENTER THE NUMBER OF COLUMNS:");

    if (!numRows || !numColumns) {
      alert("Invalid input. Please enter valid values for rows and columns.");
      return;
    }
    const body = context.document.body;

    const wrdTbl = body.insertTable(
      Number(numRows),
      Number(numColumns),
      Word.RangeLocation.end
    );

    const wrdRows = wrdTbl.rows;
    const row_1 = wrdRows.getFirst();
    row_1.shadingColor = "#D9D9D9";
    row_1.font.size = 14;
    row_1.horizontalAlignment = "Centered";
    row_1.verticalAlignment = "Center";
    await context.sync();
    const tblBorders = wrdTbl.getBorder(Word.BorderLocation.all);
    tblBorders.type = Word.BorderType.single;
    const tblBordersIV = wrdTbl.getBorder(Word.BorderLocation.insideVertical);
    tblBordersIV.type = "Single";
    const tblBordersOT = wrdTbl.getBorder(Word.BorderLocation.outside);
    tblBordersOT.type = "Single";
    wrdTbl.load("headerRowCount");
    wrdTbl.styleLastColumn = true;
    await context.sync();
    wrdRows.load("items");
    await context.sync();
    for (let i = 0; i < Number(numRows); i++) {
      wrdRows.items[i].verticalAlignment = Word.VerticalAlignment.center;
      wrdRows.items[i].horizontalAlignment = "Centered";
      wrdRows.items[i].font.name = "Times New Roman";
      wrdRows.items[i].font.size = 12;
      wrdRows.items[i].preferredHeight = 20;
      wrdRows.items[i].cells.load("items");
      await context.sync();
    }
    row_1.font.name = "Arial Bold";
  });
}
////////////////////////////////////////////////////////////////////////

////////////////////////////////////Heading/////////////////////////////
headingDropdown.addEventListener("change", () => {
  switch (headingDropdown.value) {
    case "h1":
      h1();
      break;
    case "h2":
      h2();
      break;
    case "h3":
      h3();
      break;
    case "h4":
      h4();
      break;
    case "h5":
      h5();
      break;
    case "h6":
      h6();
      break;
    case "h7":
      h7();
      break;
    case "h8":
      h8();
      break;
    case "h9":
      h9();
      break;
    case "h10":
      h10();
      break;
  }
});

async function h1() {
  //SELECTION
}
async function h2() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. GENERAL SPECIFICATIONS:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h3() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. BLOCK DIAGRAM:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h4() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. INPUT TERMINAL PIN ASSIGNMENT:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h5() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. LCD OPTICAL CHARACTERISTICS:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h6() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. ELECTRICAL CHARACTERISTICS:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h7() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. AC CHARACTERISTICS:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h8() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. CTP SPECIFICATION:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h9() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. RELIABILITY TEST RESULT:\n`,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function h10() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    count += 1;
    subCount = 0;
    const textBox = selection.insertText(
      `${count}. `,
      Word.InsertLocation.replace
    );
    textBox.font.size = 16;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
////////////////////////////////////////////////////////////////////////

/////////////////////////////////Sub-Heading/////////////////////////////
subHeadingDropdown.addEventListener("change", () => {
  switch (subHeadingDropdown.value) {
    case "sh1":
      sh1();
      break;
    case "sh2":
      sh2();
      break;
    case "sh3":
      sh3();
      break;
    case "sh4":
      sh4();
      break;
    case "sh5":
      sh5();
      break;
    case "sh6":
      sh6();
      break;
    case "sh7":
      sh7();
      break;
    case "sh8":
      sh8();
      break;
    case "sh9":
      sh9();
      break;
    case "sh10":
      sh10();
      break;
    case "sh11":
      sh11();
      break;
    case "sh12":
      sh12();
      break;
    case "sh13":
      sh13();
      break;
    case "sh14":
      sh14();
      break;
    case "sh15":
      sh15();
      break;
    case "sh16":
      sh16();
      break;
  }
});

async function sh1() {
  //SELCETION
}
async function sh2() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. CTP Features:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh3() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Mechanical Information:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh4() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. TFT PIN Define:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh5() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. CTP PIN Define:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh6() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Optical Specification:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh7() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Absolute Maximum Rating:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh8() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}. DC Electrical Characteristics:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh9() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. LED Backlight Characteristics:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh10() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Input Clock and Data Timing:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh11() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Input Timing:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh12() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. DE MODE:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh13() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}. SYNC Mode:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh14() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Data Input Format:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh15() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. Electrical Characteristics:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function sh16() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subCount += 1;
    subSubCount = 0;
    const textBox = selection.insertText(
      `${count}.${subCount}. `,
      Word.InsertLocation.end
    );
    textBox.font.size = 14;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
////////////////////////////////////////////////////////////////////////

////////////////////////////////SUB-SUB HEADING/////////////////////////
subSubHeading.addEventListener("change", () => {
  switch (subSubHeading.value) {
    case "ss1":
      ss1();
      break;
    case "ss2":
      ss2();
      break;
    case "ss3":
      ss3();
      break;
    case "ss4":
      ss4();
      break;
    case "ss5":
      ss5();
      break;
  }
});
async function ss1() {
  //SELCETION
}
async function ss2() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subSubCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}.${subSubCount}. Absolute Maximum Rating:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 12;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function ss3() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subSubCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}.${subSubCount}. DC Electrical Characteristics:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 12;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function ss4() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subSubCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}.${subSubCount}. Characteristics:\n`,
      Word.InsertLocation.end
    );
    textBox.font.size = 12;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}
async function ss5() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    subSubCount += 1;
    const textBox = selection.insertText(
      `${count}.${subCount}.${subSubCount}.`,
      Word.InsertLocation.end
    );
    textBox.font.size = 12;
    textBox.font.name = "Times New Roman Bold";
    const range = textBox.getRange("After");
    range.select();
    context.load(textBox);
    await context.sync();
  });
}

////////////////////////////////////////////////////////////////////////

/////////////////////////////////TRY-RUN/////////////////////////////

async function run() {}

////////////////////////////////////////////////////////////////////////
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
