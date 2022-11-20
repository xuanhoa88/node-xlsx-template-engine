const path = require("path");
const sizeOf = require("image-size");
const fs = require("fs");
const etree = require("elementtree");
const JSZip = require("jszip");
const assign = require("lodash/assign");
const get = require("lodash/get");
const upperCase = require("lodash/upperCase");

const DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CALC_CHAIN_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const SHARED_STRINGS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

function _toArrayBuffer(buffer) {
  const ab = new ArrayBuffer(buffer.length);
  const view = new Uint8Array(ab);
  for (let i = 0; i < buffer.length; ++i) {
    view[i] = buffer[i];
  }
  return ab;
}

// Turn a number like 27 into a reference like "AA"
function _numToChar(num) {
  let str = "";

  for (let i = 0; num > 0; ++i) {
    const remainder = num % 26;
    let charCode = remainder + 64;

    num = (num - remainder) / 26;

    // Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
    if (remainder === 0) {
      // 26 -> Z
      charCode = 90;
      --num;
    }

    str = String.fromCharCode(charCode) + str;
  }

  return str;
}

// Turn a reference like "AA" into a number like 27
function _charToNum(str) {
  let num = 0;
  for (let idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
    const char = str.charCodeAt(idx) - 64; // A -> 1; B -> 2; ... Z->26
    const multiplier = Math.pow(26, iteration);
    num += multiplier * char;
  }
  return num;
}

/**
 * Check if the buffer image is supported by the library before return it
 * @param {Buffer} buffer the final buffer image
 */
function _checkImage(buffer) {
  try {
    sizeOf(buffer);
    return buffer;
  } catch (error) {
    throw new TypeError("imageObj cannot be parse as a buffer image");
  }
}

// Turn a value of any type into a string
function _stringify(value) {
  if (value instanceof Date) {
    //In Excel date is a number of days since 01/01/1900
    //           timestamp in ms    to days      + number of days from 1900 to 1970
    return Number(value.getTime() / (1000 * 60 * 60 * 24) + 25569);
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return Number(value).toString();
  }
  if (typeof value === "string") {
    return String(value).toString();
  }

  return "";
}

/**
 * Create a new workbook. Either pass the raw data of a .xlsx file,
 * or call `loadTemplate()` later.
 */
class Workbook {
  constructor(option = {}) {
    this.archive = null;
    this.sharedStrings = [];
    this.sharedStringsLookup = {};
    this.option = assign(
      {
        moveImages: false,
        subsituteAllTableRow: false,
        moveSameLineImages: false,
        imageRatio: 100,
        pushDownPageBreakOnTableSubstitution: false,
        imageRootPath: null,
        handleImageError: null,
      },
      option
    );
    this.sharedStringsPath = "";
    this.sheets = [];
    this.sheet = null;
    this.workbook = null;
    this.workbookPath = null;
    this.contentTypes = null;
    this.prefix = null;
    this.workbookRels = null;
    this.calChainRel = null;
    this.calcChainPath = "";
  }

  /**
   * Load a .xlsx file from a byte array.
   */
  loadTemplate(data) {
    if (Buffer.isBuffer(data)) {
      data = data.toString("binary");
    } else if (fs.existsSync(data)) {
      data = fs.readFileSync(data);
    }

    this.archive = new JSZip(data, { base64: false, checkCRC32: true });

    // Load relationships
    const rels = etree
      .parse(this.archive.file("_rels/.rels").asText())
      .getroot();
    const workbookPath = rels.find(
      "Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']"
    ).attrib.Target;

    this.workbookPath = workbookPath;
    this.prefix = path.dirname(workbookPath);
    this.workbook = etree
      .parse(this.archive.file(workbookPath).asText())
      .getroot();
    this.workbookRels = etree
      .parse(
        this.archive
          .file(
            this.prefix +
              "/" +
              "_rels" +
              "/" +
              path.basename(workbookPath) +
              ".rels"
          )
          .asText()
      )
      .getroot();
    this.sheets = this.loadSheets(
      this.prefix,
      this.workbook,
      this.workbookRels
    );
    this.calChainRel = this.workbookRels.find(
      "Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']"
    );

    if (this.calChainRel) {
      this.calcChainPath = this.prefix + "/" + this.calChainRel.attrib.Target;
    }

    this.sharedStringsPath =
      this.prefix +
      "/" +
      this.workbookRels.find(
        "Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']"
      ).attrib.Target;
    this.sharedStrings = [];
    etree
      .parse(this.archive.file(this.sharedStringsPath).asText())
      .getroot()
      .findall("si")
      .forEach((si) => {
        const t = { text: "" };
        si.findall("t").forEach((tmp) => {
          t.text += tmp.text;
        });
        si.findall("r/t").forEach((tmp) => {
          t.text += tmp.text;
        });
        this.sharedStrings.push(t.text);
        this.sharedStringsLookup[t.text] = this.sharedStrings.length - 1;
      });

    this.contentTypes = etree
      .parse(this.archive.file("[Content_Types].xml").asText())
      .getroot();
    if (this.contentTypes.find('Default[@Extension="jpg"]') === null) {
      etree.SubElement(this.contentTypes, "Default", {
        ContentType: "image/png",
        Extension: "jpg",
      });
    }
  }

  /**
   * Delete unused sheets if needed
   */
  deleteSheet(sheetName) {
    const sheet = this.loadSheet(sheetName);

    const sh = this.workbook.find("sheets/sheet[@sheetId='" + sheet.id + "']");
    this.workbook.find("sheets").remove(sh);

    const rel = this.workbookRels.find(
      "Relationship[@Id='" + sh.attrib["r:id"] + "']"
    );
    this.workbookRels.remove(rel);

    this._rebuild();
    return this;
  }

  /**
   * Clone sheets in current workbook template
   */
  copySheet(sheetName, copyName, binary = true) {
    const sheet = this.loadSheet(sheetName); //filename, name , id, root
    const newSheetIndex = (
      this.workbook.findall("sheets/sheet").length + 1
    ).toString();
    const fileName = "worksheets" + "/" + "sheet" + newSheetIndex + ".xml";
    const arcName = this.prefix + "/" + fileName;
    // Copy sheet file
    this.archive.file(arcName, etree.tostring(sheet.root));
    this.archive.files[arcName].options.binary = binary;
    // copy sheet name in workbook
    const newSheet = etree.SubElement(this.workbook.find("sheets"), "sheet");
    newSheet.attrib.name = copyName || "Sheet" + newSheetIndex;
    newSheet.attrib.sheetId = newSheetIndex;
    newSheet.attrib["r:id"] = "rId" + newSheetIndex;
    // Copy definedName if any
    this.workbook.findall("definedNames/definedName").forEach((element) => {
      if (
        element.text &&
        element.text.split("!").length &&
        element.text.split("!")[0] == sheetName
      ) {
        const newDefinedName = etree.SubElement(
          this.workbook.find("definedNames"),
          "definedName",
          element.attrib
        );
        newDefinedName.text = `${copyName}!${element.text.split("!")[1]}`;
        newDefinedName.attrib.localSheetId = newSheetIndex - 1;
      }
    });

    const newRel = etree.SubElement(this.workbookRels, "Relationship");
    newRel.attrib.Type =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    newRel.attrib.Target = fileName;

    //Copy rels sheet - TODO : Maybe we can copy also the 'Target' files in rels, but Excel make this automaticly
    const relFileName =
      "worksheets" + "/_rels/" + "sheet" + newSheetIndex + ".xml.rels";
    const relArcName = this.prefix + "/" + relFileName;
    this.archive.file(
      relArcName,
      etree.tostring(this.loadSheetRels(sheet.filename).root)
    );
    this.archive.files[relArcName].options.binary = true;

    this._rebuild();
    return this;
  }

  /**
   *  Partially rebuild after copy/delete sheets
   */
  _rebuild() {
    //each <sheet> 'r:id' attribute in '\xl\workbook.xml'
    //must point to correct <Relationship> 'Id' in xl\_rels\workbook.xml.rels
    const order = ["worksheet", "theme", "styles", "sharedStrings"];

    this.workbookRels
      .findall("*")
      .sort((rel1, rel2) => {
        //using order
        const index1 = order.indexOf(path.basename(rel1.attrib.Type));
        const index2 = order.indexOf(path.basename(rel2.attrib.Type));
        // If the attrib.Type is not in the order list, go to the end of sort
        // Maybe we can do it more gracefully with the boolean operator
        if (index1 < 0 && index2 >= 0) return 1; // rel1 go after rel2
        if (index1 >= 0 && index2 < 0) return -1; // rel1 go before rel2
        if (index1 < 0 && index2 < 0) return 0; // change nothing

        if (index1 + index2 == 0) {
          if (rel1.attrib.Id && rel2.attrib.Id)
            return rel1.attrib.Id.substring(3) - rel2.attrib.Id.substring(3);
          return rel1._id - rel2._id;
        }
        return index1 - index2;
      })
      .forEach((item, index) => {
        item.attrib.Id = "rId" + (index + 1);
      });

    this.workbook.findall("sheets/sheet").forEach((item, index) => {
      item.attrib["r:id"] = "rId" + (index + 1);
      item.attrib.sheetId = (index + 1).toString();
    });

    this.archive.file(
      this.prefix +
        "/" +
        "_rels" +
        "/" +
        path.basename(this.workbookPath) +
        ".rels",
      etree.tostring(this.workbookRels)
    );
    this.archive.file(this.workbookPath, etree.tostring(this.workbook));
    this.sheets = this.loadSheets(
      this.prefix,
      this.workbook,
      this.workbookRels
    );
  }

  /**
   * Interpolate values for all the sheets using the given substitutions
   * (an object).
   */
  substituteAll(substitutions) {
    const sheets = this.loadSheets(
      this.prefix,
      this.workbook,
      this.workbookRels
    );
    sheets.forEach((sheet) => {
      this.substitute(sheet.id, substitutions);
    });
  }

  /**
   * Interpolate values for the sheet with the given number (1-based) or
   * name (if a string) using the given substitutions (an object).
   */
  substitute(sheetName, substitutions) {
    const sheet = this.loadSheet(sheetName);
    this.sheet = sheet;

    const dimension = sheet.root.find("dimension");
    const sheetData = sheet.root.find("sheetData");
    const namedTables = this.loadTables(sheet.root, sheet.filename);
    const rows = [];

    let currentRow = null;
    let totalRowsInserted = 0;
    let totalColumnsInserted = 0;
    let drawing = null;

    const rels = this.loadSheetRels(sheet.filename);
    sheetData.findall("row").forEach((row) => {
      row.attrib.r = currentRow = this.getCurrentRow(row, totalRowsInserted);
      rows.push(row);

      const cells = [];
      const newTableRows = [];
      const cellsForsubstituteTable = []; // Contains all the row cells when substitute tables

      let cellsInserted = 0;

      row.findall("c").forEach((cell) => {
        let appendCell = true;
        cell.attrib.r = this.getCurrentCell(cell, currentRow, cellsInserted);

        // If c[@t="s"] (string column), look up /c/v@text as integer in
        // `this.sharedStrings`
        if (cell.attrib.t === "s") {
          // Look for a shared string that may contain placeholders
          const cellValue = cell.find("v");
          const stringIndex = parseInt(cellValue.text, 10);
          let string = this.sharedStrings[stringIndex];

          if (string === undefined) {
            return;
          }

          // Loop over placeholders
          this.extractPlaceholders(string).forEach((placeholder) => {
            // Only substitute things for which we have a substitution
            let substitution = get(substitutions, placeholder.name, "");
            let newCellsInserted = 0;

            if (
              placeholder.full &&
              placeholder.type === "table" &&
              substitution instanceof Array
            ) {
              if (placeholder.subType === "image" && drawing == null) {
                if (rels) {
                  drawing = this.loadDrawing(
                    sheet.root,
                    sheet.filename,
                    rels.root
                  );
                } else {
                  console.log(
                    "Need to implement initRels. Or init this with Excel"
                  );
                }
              }
              cellsForsubstituteTable.push(cell); // When substitute table, push (all) the cell
              newCellsInserted = this.substituteTable(
                row,
                newTableRows,
                cells,
                cell,
                namedTables,
                substitution,
                placeholder.key,
                placeholder,
                drawing
              );

              // don't double-insert cells
              // this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
              if (newCellsInserted !== 0 || substitution.length) {
                if (substitution.length === 1) {
                  appendCell = true;
                }
                if (substitution[0][placeholder.key] instanceof Array) {
                  appendCell = false;
                }
              }

              // Did we insert new columns (array values)?
              if (newCellsInserted !== 0) {
                cellsInserted += newCellsInserted;
                this.pushRight(
                  this.workbook,
                  sheet.root,
                  cell.attrib.r,
                  newCellsInserted
                );
              }
            } else if (
              placeholder.full &&
              placeholder.type === "normal" &&
              substitution instanceof Array
            ) {
              appendCell = false; // don't double-insert cells
              newCellsInserted = this.substituteArray(
                cells,
                cell,
                substitution
              );

              if (newCellsInserted !== 0) {
                cellsInserted += newCellsInserted;
                this.pushRight(
                  this.workbook,
                  sheet.root,
                  cell.attrib.r,
                  newCellsInserted
                );
              }
            } else if (placeholder.type === "image" && placeholder.full) {
              if (rels != null) {
                if (drawing == null) {
                  drawing = this.loadDrawing(
                    sheet.root,
                    sheet.filename,
                    rels.root
                  );
                }
                string = this.substituteImage(
                  cell,
                  string,
                  placeholder,
                  substitution,
                  drawing
                );
              } else {
                console.log(
                  "Need to implement initRels. Or init this with Excel"
                );
              }
            } else {
              if (placeholder.key) {
                substitution = get(
                  substitutions,
                  placeholder.name + "." + placeholder.key
                );
              }
              string = this.substituteScalar(
                cell,
                string,
                placeholder,
                substitution
              );
            }
          });
        }

        // if we are inserting columns, we may not want to keep the original cell anymore
        if (appendCell) {
          cells.push(cell);
        }
      }); // cells loop

      // We may have inserted columns, so re-build the children of the row
      this.replaceChildren(row, cells);

      // Update row spans attribute
      if (cellsInserted !== 0) {
        this.updateRowSpan(row, cellsInserted);

        if (cellsInserted > totalColumnsInserted) {
          totalColumnsInserted = cellsInserted;
        }
      }

      // Add newly inserted rows
      if (newTableRows.length > 0) {
        // Move images for each subsitute array if option is active
        if (this.option["moveImages"] && rels) {
          if (drawing == null) {
            // Maybe we can load drawing at the begining of function and remove all the this.loadDrawing() along the function ?
            // If we make this, we create all the time the drawing file (like rels file at this moment)
            drawing = this.loadDrawing(sheet.root, sheet.filename, rels.root);
          }
          if (drawing != null) {
            this.moveAllImages(drawing, row.attrib.r, newTableRows.length);
          }
        }

        // Filter all the cellsForsubstituteTable cell with the 'row' cell
        const cellsOverTable = row
          .findall("c")
          .filter((cell) => !cellsForsubstituteTable.includes(cell));

        newTableRows.forEach((row) => {
          if (this.option && this.option.subsituteAllTableRow) {
            // I happend the other cell in substitute new table rows
            cellsOverTable.forEach((cellOverTable) => {
              const newCell = this.cloneElement(cellOverTable);
              newCell.attrib.r = this.joinRef({
                row: row.attrib.r,
                col: this.splitRef(newCell.attrib.r).col,
              });
              row.append(newCell);
            });
            // I sort the cell in the new row
            const newSortRow = row.findall("c").sort((a, b) => {
              const colA = this.splitRef(a.attrib.r).col;
              const colB = this.splitRef(b.attrib.r).col;
              return _charToNum(colA) - _charToNum(colB);
            });
            // And I replace the cell
            this.replaceChildren(row, newSortRow);
          }

          rows.push(row);
          ++totalRowsInserted;
        });
        this.pushDown(
          this.workbook,
          sheet.root,
          namedTables,
          currentRow,
          newTableRows.length
        );
      }
    }); // rows loop

    // We may have inserted rows, so re-build the children of the sheetData
    this.replaceChildren(sheetData, rows);

    // Update placeholders in table column headers
    this.substituteTableColumnHeaders(namedTables, substitutions);

    // Update placeholders in hyperlinks
    this.substituteHyperlinks(rels, substitutions);

    // Update <dimension /> if we added rows or columns
    if (dimension) {
      if (totalRowsInserted > 0 || totalColumnsInserted > 0) {
        const dimensionRange = this.splitRange(dimension.attrib.ref);
        const dimensionEndRef = this.splitRef(dimensionRange.end);

        dimensionEndRef.row += totalRowsInserted;
        dimensionEndRef.col = _numToChar(
          _charToNum(dimensionEndRef.col) + totalColumnsInserted
        );
        dimensionRange.end = this.joinRef(dimensionEndRef);

        dimension.attrib.ref = this.joinRange(dimensionRange);
      }
    }

    //Here we are forcing the values in formulas to be recalculated
    // existing as well as just substituted
    sheetData.findall("row").forEach((row) => {
      row.findall("c").forEach((cell) => {
        const formulas = cell.findall("f");
        if (formulas && formulas.length > 0) {
          cell.findall("v").forEach((v) => {
            cell.remove(v);
          });
        }
      });
    });

    // Write back the modified XML trees
    this.archive.file(sheet.filename, etree.tostring(sheet.root));
    this.archive.file(this.workbookPath, etree.tostring(this.workbook));
    if (rels) {
      this.archive.file(rels.filename, etree.tostring(rels.root));
    }
    this.archive.file("[Content_Types].xml", etree.tostring(this.contentTypes));
    // Remove calc chain - Excel will re-build, and we may have moved some formulae
    if (this.calcChainPath && this.archive.file(this.calcChainPath)) {
      this.archive.remove(this.calcChainPath);
    }

    this.writeSharedStrings();
    this.writeTables(namedTables);
    this.writeDrawing(drawing);
  }

  /**
   * Generate a new binary .xlsx file
   */
  generate(options) {
    if (!options) {
      options = {
        base64: false,
      };
    }

    return this.archive.generate(options);
  }

  // Write back the new shared strings list
  writeSharedStrings() {
    const root = etree
      .parse(this.archive.file(this.sharedStringsPath).asText())
      .getroot();

    root.delSlice(0, root.getchildren().length);

    this.sharedStrings.forEach((string) => {
      const si = new etree.Element("si");
      const t = new etree.Element("t");

      t.text = string;
      si.append(t);
      root.append(si);
    });

    root.attrib.count = this.sharedStrings.length;
    root.attrib.uniqueCount = this.sharedStrings.length;

    this.archive.file(this.sharedStringsPath, etree.tostring(root));
  }

  // Add a new shared string
  addSharedString(s) {
    const idx = this.sharedStrings.length;
    this.sharedStrings.push(s);
    this.sharedStringsLookup[s] = idx;

    return idx;
  }

  // Get the number of a shared string, adding a new one if necessary.
  stringIndex(s) {
    let idx = this.sharedStringsLookup[s];
    if (idx === undefined) {
      idx = this.addSharedString(s);
    }
    return idx;
  }

  // Replace a shared string with a new one at the same index. Return the
  // index.
  replaceString(oldString, newString) {
    let idx = this.sharedStringsLookup[oldString];
    if (idx === undefined) {
      idx = this.addSharedString(newString);
    } else {
      this.sharedStrings[idx] = newString;
      delete this.sharedStringsLookup[oldString];
      this.sharedStringsLookup[newString] = idx;
    }

    return idx;
  }

  // Get a list of sheet ids, names and filenames
  loadSheets(prefix, workbook, workbookRels) {
    const sheets = [];
    workbook.findall("sheets/sheet").forEach((sheet) => {
      const sheetId = sheet.attrib.sheetId;
      const relId = sheet.attrib["r:id"];
      const relationship = workbookRels.find(
        "Relationship[@Id='" + relId + "']"
      );
      const filename = prefix + "/" + relationship.attrib.Target;

      sheets.push({
        id: parseInt(sheetId, 10),
        name: sheet.attrib.name,
        filename,
      });
    });

    return sheets;
  }

  // Get sheet a sheet, including filename and name
  loadSheet(sheet) {
    let info = null;

    for (let i = 0; i < this.sheets.length; ++i) {
      if (
        (typeof sheet === "number" && this.sheets[i].id === sheet) ||
        this.sheets[i].name === sheet
      ) {
        info = this.sheets[i];
        break;
      }
    }

    if (info === null && typeof sheet === "number") {
      //Get the sheet that corresponds to the 0 based index if the id does not work
      info = this.sheets[sheet - 1];
    }

    if (info === null) {
      throw new Error("Sheet " + sheet + " not found");
    }

    return {
      filename: info.filename,
      name: info.name,
      id: info.id,
      root: etree.parse(this.archive.file(info.filename).asText()).getroot(),
    };
  }

  //Load rels for a sheetName
  loadSheetRels(sheetFilename) {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    const relsFilename = path
        .join(sheetDirectory, "_rels", sheetName + ".rels")
        .replace(/\\/g, "/"),
      relsFile = this.archive.file(relsFilename);
    if (relsFile === null) {
      return this.initSheetRels(sheetFilename);
    }
    return {
      filename: relsFilename,
      root: etree.parse(relsFile.asText()).getroot(),
    };
  }

  initSheetRels(sheetFilename) {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    const relsFilename = path
      .join(sheetDirectory, "_rels", sheetName + ".rels")
      .replace(/\\/g, "/");
    const element = etree.Element;
    const ElementTree = etree.ElementTree;
    const root = element("Relationships");
    root.set(
      "xmlns",
      "http://schemas.openxmlformats.org/package/2006/relationships"
    );
    const relsEtree = new ElementTree(root);
    return { filename: relsFilename, root: relsEtree.getroot() };
  }

  // Load Drawing file
  loadDrawing(sheet, sheetFilename, rels) {
    const drawingPart = sheet.find("drawing");
    if (drawingPart === null) {
      return this.initDrawing(sheet, rels);
    }
    const sheetDirectory = path.dirname(sheetFilename);
    const drawing = { filename: "", root: null };
    const relationshipId = drawingPart.attrib["r:id"];
    const target = rels.find("Relationship[@Id='" + relationshipId + "']")
      .attrib.Target;
    const drawingFilename = path
      .join(sheetDirectory, target)
      .replace(/\\/g, "/");
    const drawingTree = etree.parse(
      this.archive.file(drawingFilename).asText()
    );
    drawing.filename = drawingFilename;
    drawing.root = drawingTree.getroot();
    drawing.relFilename =
      path.dirname(drawingFilename) +
      "/_rels/" +
      path.basename(drawingFilename) +
      ".rels";
    drawing.relRoot = etree
      .parse(this.archive.file(drawing.relFilename).asText())
      .getroot();
    return drawing;
  }

  addContentType(partName, contentType) {
    etree.SubElement(this.contentTypes, "Override", {
      ContentType: contentType,
      PartName: partName,
    });
  }

  initDrawing(sheet, rels) {
    const maxId = this.findMaxId(rels, "Relationship", "Id", /rId(\d*)/);
    const rel = etree.SubElement(rels, "Relationship");
    sheet.insert(
      sheet._children.length,
      etree.Element("drawing", { "r:id": "rId" + maxId })
    );
    rel.set("Id", "rId" + maxId);
    rel.set(
      "Type",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    );
    const drawing = {};
    const drawingFilename =
      "drawing" +
      this.findMaxFileId(/xl\/drawings\/drawing\d*\.xml/, /drawing(\d*)\.xml/) +
      ".xml";
    rel.set("Target", "../drawings/" + drawingFilename);
    drawing.root = etree.Element("xdr:wsDr");
    drawing.root.set(
      "xmlns:xdr",
      "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    );
    drawing.root.set(
      "xmlns:a",
      "http://schemas.openxmlformats.org/drawingml/2006/main"
    );
    drawing.filename = "xl/drawings/" + drawingFilename;
    drawing.relFilename = "xl/drawings/_rels/" + drawingFilename + ".rels";
    drawing.relRoot = etree.Element("Relationships");
    drawing.relRoot.set(
      "xmlns",
      "http://schemas.openxmlformats.org/package/2006/relationships"
    );
    this.addContentType(
      "/" + drawing.filename,
      "application/vnd.openxmlformats-officedocument.drawing+xml"
    );
    return drawing;
  }

  // Write Drawing file
  writeDrawing(drawing) {
    if (drawing !== null) {
      this.archive.file(drawing.filename, etree.tostring(drawing.root));
      this.archive.file(drawing.relFilename, etree.tostring(drawing.relRoot));
    }
  }

  // Move all images after fromRow of nbRow row
  moveAllImages(drawing, fromRow, nbRow) {
    drawing.root.getchildren().forEach((drawElement) => {
      if (drawElement.tag == "xdr:twoCellAnchor") {
        this._moveTwoCellAnchor(drawElement, fromRow, nbRow);
      }
      // TODO : make the other tags image
    });
  }

  _moveImage(drawingElement, fromRow, nbRow) {
    drawingElement.find("xdr:from").find("xdr:row").text =
      Number.parseInt(
        drawingElement.find("xdr:from").find("xdr:row").text,
        10
      ) + Number.parseInt(nbRow, 10);
    drawingElement.find("xdr:to").find("xdr:row").text =
      Number.parseInt(drawingElement.find("xdr:to").find("xdr:row").text, 10) +
      Number.parseInt(nbRow, 10);
  }

  // Move TwoCellAnchor tag images after fromRow of nbRow row
  _moveTwoCellAnchor(drawingElement, fromRow, nbRow) {
    if (this.option["moveSameLineImages"]) {
      if (
        parseInt(drawingElement.find("xdr:from").find("xdr:row").text) + 1 >=
        fromRow
      ) {
        this._moveImage(drawingElement, fromRow, nbRow);
      }
    } else {
      if (
        parseInt(drawingElement.find("xdr:from").find("xdr:row").text) + 1 >
        fromRow
      ) {
        this._moveImage(drawingElement, fromRow, nbRow);
      }
    }
  }

  // Load tables for a given sheet
  loadTables(sheet, sheetFilename) {
    const relsFile = this.archive.file(
      path.dirname(sheetFilename) +
        "/" +
        "_rels" +
        "/" +
        path.basename(sheetFilename) +
        ".rels"
    );
    const tables = [];

    if (relsFile === null) {
      return tables;
    }

    const rels = etree.parse(relsFile.asText()).getroot();

    sheet.findall("tableParts/tablePart").forEach((tablePart) => {
      const relationshipId = tablePart.attrib["r:id"];
      const target = rels.find("Relationship[@Id='" + relationshipId + "']")
        .attrib.Target;
      const tableFilename = target.replace("..", this.prefix);
      const tableTree = etree.parse(this.archive.file(tableFilename).asText());

      tables.push({
        filename: tableFilename,
        root: tableTree.getroot(),
      });
    });

    return tables;
  }

  // Write back possibly-modified tables
  writeTables(tables) {
    tables.forEach((namedTable) => {
      this.archive.file(namedTable.filename, etree.tostring(namedTable.root));
    });
  }

  //Perform substitution in hyperlinks
  substituteHyperlinks(rels, substitutions) {
    etree.parse(this.archive.file(this.sharedStringsPath).asText()).getroot();
    if (rels === null) {
      return;
    }
    const relationships = rels.root._children;
    relationships.forEach((relationship) => {
      if (relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {
        let target = relationship.attrib.Target;

        //Double-decode due to excel double encoding url placeholders
        target = decodeURI(decodeURI(target));
        this.extractPlaceholders(target).forEach((placeholder) => {
          const substitution = substitutions[placeholder.name];

          if (substitution === undefined) {
            return;
          }
          target = target.replace(
            placeholder.placeholder,
            _stringify(substitution)
          );

          relationship.attrib.Target = encodeURI(target);
        });
      }
    });
  }

  // Perform substitution in table headers
  substituteTableColumnHeaders(tables, substitutions) {
    tables.forEach((table) => {
      const root = table.root;
      const columns = root.find("tableColumns");
      let autoFilter = root.find("autoFilter");
      let tableRange = this.splitRange(root.attrib.ref);
      let idx = 0;
      let inserted = 0;
      const newColumns = [];

      columns.findall("tableColumn").forEach((col) => {
        ++idx;
        col.attrib.id = Number(idx).toString();
        newColumns.push(col);

        const name = col.attrib.name;

        this.extractPlaceholders(name).forEach((placeholder) => {
          const substitution = substitutions[placeholder.name];
          if (substitution === undefined) {
            return;
          }

          // Array -> new columns
          if (
            placeholder.full &&
            placeholder.type === "normal" &&
            substitution instanceof Array
          ) {
            substitution.forEach((element, i) => {
              let newCol = col;
              if (i > 0) {
                newCol = this.cloneElement(newCol);
                newCol.attrib.id = Number(++idx).toString();
                newColumns.push(newCol);
                ++inserted;
                tableRange.end = this.nextCol(tableRange.end);
              }
              newCol.attrib.name = _stringify(element);
            });
            // Normal placeholder
          } else {
            name = name.replace(
              placeholder.placeholder,
              _stringify(substitution)
            );
            col.attrib.name = name;
          }
        });
      });

      this.replaceChildren(columns, newColumns);

      // Update range if we inserted columns
      if (inserted > 0) {
        columns.attrib.count = Number(idx).toString();
        root.attrib.ref = this.joinRange(tableRange);
        if (autoFilter !== null) {
          // XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = this.joinRange(tableRange);
        }
      }

      //update ranges for totalsRowCount
      const tableRoot = table.root;
      const tableStart = this.splitRef(tableRange.start);
      const tableEnd = this.splitRef(tableRange.end);

      tableRange = this.splitRange(tableRoot.attrib.ref);

      if (tableRoot.attrib.totalsRowCount) {
        autoFilter = tableRoot.find("autoFilter");
        if (autoFilter !== null) {
          autoFilter.attrib.ref = this.joinRange({
            start: this.joinRef(tableStart),
            end: this.joinRef(tableEnd),
          });
        }

        ++tableEnd.row;
        tableRoot.attrib.ref = this.joinRange({
          start: this.joinRef(tableStart),
          end: this.joinRef(tableEnd),
        });
      }
    });
  }

  // Return a list of tokens that may exist in the string.
  // Keys are: `placeholder` (the full placeholder, including the `${}`
  // delineators), `name` (the name part of the token), `key` (the object key
  // for `table` tokens), `full` (boolean indicating whether this placeholder
  // is the entirety of the string) and `type` (one of `table` or `cell`)
  extractPlaceholders(string) {
    // Yes, that's right. It's a bunch of brackets and question marks and stuff.
    const re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?(?::(.+?))??}/g;
    const matches = [];
    let match = null;
    while ((match = re.exec(string)) !== null) {
      matches.push({
        placeholder: match[0],
        type: match[1] || "normal",
        name: match[2],
        key: match[3],
        subType: match[4],
        full: match[0].length === string.length,
      });
    }

    return matches;
  }

  // Split a reference into an object with keys `row` and `col` and,
  // optionally, `table`, `rowAbsolute` and `colAbsolute`.
  splitRef(ref) {
    const match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)?(\$)?([0-9]+)/);
    return {
      table: (match && match[1]) || null,
      colAbsolute: Boolean(match && match[2]),
      col: (match && match[3]) || "",
      rowAbsolute: Boolean(match && match[4]),
      row: parseInt(match && match[5], 10),
    };
  }

  // Join an object with keys `row` and `col` into a single reference string
  joinRef(ref) {
    return (
      (ref.table ? ref.table + "!" : "") +
      (ref.colAbsolute ? "$" : "") +
      upperCase(ref.col) +
      (ref.rowAbsolute ? "$" : "") +
      Number(ref.row).toString()
    );
  }

  // Get the next column's cell reference given a reference like "B2".
  nextCol(ref) {
    return upperCase(ref).replace(/[A-Z]+/, (match) =>
      _numToChar(_charToNum(match) + 1)
    );
  }

  // Get the next row's cell reference given a reference like "B2".
  nextRow(ref) {
    return upperCase(ref).replace(/[0-9]+/, (match) =>
      (parseInt(match, 10) + 1).toString()
    );
  }

  // Is ref a range?
  isRange(ref) {
    return ref.indexOf(":") !== -1;
  }

  // Is ref inside the table defined by startRef and endRef?
  isWithin(ref, startRef, endRef) {
    const start = this.splitRef(startRef);
    const end = this.splitRef(endRef);
    const target = this.splitRef(ref);

    start.col = _charToNum(start.col);
    end.col = _charToNum(end.col);
    target.col = _charToNum(target.col);

    return (
      start.row <= target.row &&
      target.row <= end.row &&
      start.col <= target.col &&
      target.col <= end.col
    );
  }

  // Insert a substitution value into a cell (c tag)
  insertCellValue(cell, substitution) {
    const cellValue = cell.find("v");
    const stringified = _stringify(substitution);

    if (typeof substitution === "string" && substitution[0] === "=") {
      //substitution, started with '=' is a formula substitution
      const formula = new etree.Element("f");
      formula.text = substitution.substr(1);
      cell.insert(1, formula);
      delete cell.attrib.t; //cellValue will be deleted later
      return formula.text;
    }

    if (typeof substitution === "number" || substitution instanceof Date) {
      delete cell.attrib.t;
      cellValue.text = stringified;
    } else if (typeof substitution === "boolean") {
      cell.attrib.t = "b";
      cellValue.text = stringified;
    } else {
      cell.attrib.t = "s";
      cellValue.text = Number(this.stringIndex(stringified)).toString();
    }

    return stringified;
  }

  // Perform substitution of a single value
  substituteScalar(cell, string, placeholder, substitution) {
    if (placeholder.full) {
      return this.insertCellValue(cell, substitution);
    }
    const newString = string.replace(
      placeholder.placeholder,
      _stringify(substitution)
    );
    cell.attrib.t = "s";
    return this.insertCellValue(cell, newString);
  }

  // Perform a columns substitution from an array
  substituteArray(cells, cell, substitution) {
    let newCellsInserted = -1; // we technically delete one before we start adding back
    let currentCell = cell.attrib.r;

    // add a cell for each element in the list
    substitution.forEach((element) => {
      ++newCellsInserted;

      if (newCellsInserted > 0) {
        currentCell = this.nextCol(currentCell);
      }

      const newCell = this.cloneElement(cell);
      this.insertCellValue(newCell, element);

      newCell.attrib.r = currentCell;
      cells.push(newCell);
    });

    return newCellsInserted;
  }

  // Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
  // Returns total number of new cells inserted on the original row.
  substituteTable(
    row,
    newTableRows,
    cells,
    cell,
    namedTables,
    substitution,
    key,
    placeholder,
    drawing
  ) {
    let newCellsInserted = 0; // on the original row

    // if no elements, blank the cell, but don't delete it
    if (substitution.length === 0) {
      delete cell.attrib.t;
      this.replaceChildren(cell, []);
    } else {
      const parentTables = namedTables.filter((namedTable) => {
        const range = this.splitRange(namedTable.root.attrib.ref);
        return this.isWithin(cell.attrib.r, range.start, range.end);
      });

      substitution.forEach((element, idx) => {
        let newRow;
        let newCell;
        let newCellsInsertedOnNewRow = 0;
        const newCells = [];
        const value = get(element, key, "");

        if (idx === 0) {
          // insert in the row where the placeholders are
          if (value instanceof Array) {
            newCellsInserted = this.substituteArray(cells, cell, value);
          } else if (placeholder.subType == "image" && value != "") {
            this.substituteImage(
              cell,
              placeholder.placeholder,
              placeholder,
              value,
              drawing
            );
          } else {
            this.insertCellValue(cell, value);
          }
        } else {
          // insert new rows (or reuse rows just inserted)

          // Do we have an existing row to use? If not, create one.
          if (idx - 1 < newTableRows.length) {
            newRow = newTableRows[idx - 1];
          } else {
            newRow = this.cloneElement(row, false);
            newRow.attrib.r = this.getCurrentRow(row, newTableRows.length + 1);
            newTableRows.push(newRow);
          }

          // Create a new cell
          newCell = this.cloneElement(cell);
          newCell.attrib.r = this.joinRef({
            row: newRow.attrib.r,
            col: this.splitRef(newCell.attrib.r).col,
          });

          if (value instanceof Array) {
            newCellsInsertedOnNewRow = this.substituteArray(
              newCells,
              newCell,
              value
            );

            // Add each of the new cells created by substituteArray()
            newCells.forEach((newCell) => {
              newRow.append(newCell);
            });

            this.updateRowSpan(newRow, newCellsInsertedOnNewRow);
          } else if (placeholder.subType == "image" && value != "") {
            this.substituteImage(
              newCell,
              placeholder.placeholder,
              placeholder,
              value,
              drawing
            );
          } else {
            this.insertCellValue(newCell, value);

            // Add the cell that previously held the placeholder
            newRow.append(newCell);
          }

          // expand named table range if necessary
          parentTables.forEach((namedTable) => {
            const tableRoot = namedTable.root;
            const autoFilter = tableRoot.find("autoFilter");
            const range = this.splitRange(tableRoot.attrib.ref);

            if (!this.isWithin(newCell.attrib.r, range.start, range.end)) {
              range.end = this.nextRow(range.end);
              tableRoot.attrib.ref = this.joinRange(range);
              if (autoFilter !== null) {
                // XXX: This is a simplification that may stomp on some configurations
                autoFilter.attrib.ref = tableRoot.attrib.ref;
              }
            }
          });
        }
      });
    }

    return newCellsInserted;
  }

  substituteImage(cell, string, placeholder, substitution, drawing) {
    this.substituteScalar(cell, string, placeholder, "");
    if (substitution == null || substitution == "") {
      // TODO : @kant2002 if image is null or empty string in user substitution data, throw an error or not ?
      // If yes, remove this test.
      return true;
    }
    // get max refid
    // update rel file.
    const maxId = this.findMaxId(
      drawing.relRoot,
      "Relationship",
      "Id",
      /rId(\d*)/
    );
    const maxFildId = this.findMaxFileId(
      /xl\/media\/image\d*.jpg/,
      /image(\d*)\.jpg/
    );
    const rel = etree.SubElement(drawing.relRoot, "Relationship");
    rel.set("Id", "rId" + maxId);
    rel.set(
      "Type",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    );

    rel.set("Target", "../media/image" + maxFildId + ".jpg");

    try {
      substitution = this.imageToBuffer(substitution);
    } catch (error) {
      if (
        this.option &&
        this.option.handleImageError &&
        typeof this.option.handleImageError === "function"
      ) {
        this.option.handleImageError(substitution, error);
      } else {
        throw error;
      }
    }

    // put image to media.
    this.archive.file(
      "xl/media/image" + maxFildId + ".jpg",
      _toArrayBuffer(substitution),
      { binary: true, base64: false }
    );
    const dimension = sizeOf(substitution);
    let imageWidth = this.pixelsToEMUs(dimension.width);
    let imageHeight = this.pixelsToEMUs(dimension.height);
    let imageInMergeCell = false;
    this.sheet.root.findall("mergeCells/mergeCell").forEach((mergeCell) => {
      // If image is in merge cell, fit the image
      if (this.cellInMergeCells(cell, mergeCell)) {
        const mergeCellWidth = this.getWidthMergeCell(mergeCell, this.sheet);
        const mergeCellHeight = this.getHeightMergeCell(mergeCell, this.sheet);
        const mergeWidthEmus = this.columnWidthToEMUs(mergeCellWidth);
        const mergeHeightEmus = this.rowHeightToEMUs(mergeCellHeight);
        // Maybe we can add an option for fit image to mergecell if image is more little. Not by default

        let widthRate = imageWidth / mergeWidthEmus;
        let heightRate = imageHeight / mergeHeightEmus;
        if (widthRate > heightRate) {
          imageWidth = Math.floor(imageWidth / widthRate);
          imageHeight = Math.floor(imageHeight / widthRate);
        } else {
          imageWidth = Math.floor(imageWidth / heightRate);
          imageHeight = Math.floor(imageHeight / heightRate);
        }
        imageInMergeCell = true;
      }
    });
    if (imageInMergeCell == false) {
      let ratio = 100;
      if (this.option && this.option.imageRatio) {
        ratio = this.option.imageRatio;
      }
      if (ratio <= 0) {
        ratio = 100;
      }
      imageWidth = Math.floor((imageWidth * ratio) / 100);
      imageHeight = Math.floor((imageHeight * ratio) / 100);
    }
    const imagePart = etree.SubElement(drawing.root, "xdr:oneCellAnchor");
    const fromPart = etree.SubElement(imagePart, "xdr:from");
    const fromCol = etree.SubElement(fromPart, "xdr:col");
    fromCol.text = (
      _charToNum(this.splitRef(cell.attrib.r).col) - 1
    ).toString();
    const fromColOff = etree.SubElement(fromPart, "xdr:colOff");
    fromColOff.text = "0";
    const fromRow = etree.SubElement(fromPart, "xdr:row");
    fromRow.text = (this.splitRef(cell.attrib.r).row - 1).toString();
    const fromRowOff = etree.SubElement(fromPart, "xdr:rowOff");
    fromRowOff.text = "0";
    const extImagePart = etree.SubElement(imagePart, "xdr:ext", {
      cx: imageWidth,
      cy: imageHeight,
    });
    const picNode = etree.SubElement(imagePart, "xdr:pic");
    const nvPicPr = etree.SubElement(picNode, "xdr:nvPicPr");
    const cNvPr = etree.SubElement(nvPicPr, "xdr:cNvPr", {
      id: maxId,
      name: "image_" + maxId,
      descr: "",
    });
    const cNvPicPr = etree.SubElement(nvPicPr, "xdr:cNvPicPr");
    const picLocks = etree.SubElement(cNvPicPr, "a:picLocks", {
      noChangeAspect: "1",
    });
    const blipFill = etree.SubElement(picNode, "xdr:blipFill");
    const blip = etree.SubElement(blipFill, "a:blip", {
      "xmlns:r":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "r:embed": "rId" + maxId,
    });
    const stretch = etree.SubElement(blipFill, "a:stretch");
    const fillRect = etree.SubElement(stretch, "a:fillRect");
    const spPr = etree.SubElement(picNode, "xdr:spPr");
    const xfrm = etree.SubElement(spPr, "a:xfrm");
    const off = etree.SubElement(xfrm, "a:off", { x: "0", y: "0" });
    const ext = etree.SubElement(xfrm, "a:ext", {
      cx: imageWidth,
      cy: imageHeight,
    });
    const prstGeom = etree.SubElement(spPr, "a:prstGeom", { prst: "rect" });
    const avLst = etree.SubElement(prstGeom, "a:avLst");
    const clientData = etree.SubElement(imagePart, "xdr:clientData");
    return true;
  }

  // Clone an element. If `deep` is true, recursively clone children
  cloneElement(element, deep) {
    const newElement = etree.Element(element.tag, element.attrib);
    newElement.text = element.text;
    newElement.tail = element.tail;

    if (deep !== false) {
      element.getchildren().forEach((child) => {
        newElement.append(this.cloneElement(child, deep));
      });
    }

    return newElement;
  }

  // Replace all children of `parent` with the nodes in the list `children`
  replaceChildren(parent, children) {
    parent.delSlice(0, parent.len());
    children.forEach((child) => {
      parent.append(child);
    });
  }

  // Calculate the current row based on a source row and a number of new rows
  // that have been inserted above
  getCurrentRow(row, rowsInserted) {
    return parseInt(row.attrib.r, 10) + rowsInserted;
  }

  // Calculate the current cell based on asource cell, the current row index,
  // and a number of new cells that have been inserted so far
  getCurrentCell(cell, currentRow, cellsInserted) {
    const colRef = this.splitRef(cell.attrib.r).col;
    const colNum = _charToNum(colRef);

    return this.joinRef({
      row: currentRow,
      col: _numToChar(colNum + cellsInserted),
    });
  }

  // Adjust the row `spans` attribute by `cellsInserted`
  updateRowSpan(row, cellsInserted) {
    if (cellsInserted !== 0 && row.attrib.spans) {
      const rowSpan = row.attrib.spans.split(":").map((f) => parseInt(f, 10));
      rowSpan[1] += cellsInserted;
      row.attrib.spans = rowSpan.join(":");
    }
  }

  // Split a range like "A1:B1" into {start: "A1", end: "B1"}
  splitRange(range) {
    const split = range.split(":");
    return {
      start: split[0],
      end: split[1],
    };
  }

  // Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}
  joinRange(range) {
    return range.start + ":" + range.end;
  }

  // Look for any merged cell or named range definitions to the right of
  // `currentCell` and push right by `numCols`.
  pushRight(workbook, sheet, currentCell, numCols) {
    const cellRef = this.splitRef(currentCell);
    const currentRow = cellRef.row;
    const currentCol = _charToNum(cellRef.col);

    // Update merged cells on the same row, at a higher column
    sheet.findall("mergeCells/mergeCell").forEach((mergeCell) => {
      const mergeRange = this.splitRange(mergeCell.attrib.ref);
      const mergeStart = this.splitRef(mergeRange.start);
      const mergeStartCol = _charToNum(mergeStart.col);
      const mergeEnd = this.splitRef(mergeRange.end);
      const mergeEndCol = _charToNum(mergeEnd.col);

      if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
        mergeStart.col = _numToChar(mergeStartCol + numCols);
        mergeEnd.col = _numToChar(mergeEndCol + numCols);

        mergeCell.attrib.ref = this.joinRange({
          start: this.joinRef(mergeStart),
          end: this.joinRef(mergeEnd),
        });
      }
    });

    // Named cells/ranges
    workbook.findall("definedNames/definedName").forEach((name) => {
      const ref = name.text;

      if (this.isRange(ref)) {
        const namedRange = this.splitRange(ref);
        const namedStart = this.splitRef(namedRange.start);
        const namedStartCol = _charToNum(namedStart.col);
        const namedEnd = this.splitRef(namedRange.end);
        const namedEndCol = _charToNum(namedEnd.col);

        if (namedStart.row === currentRow && currentCol < namedStartCol) {
          namedStart.col = _numToChar(namedStartCol + numCols);
          namedEnd.col = _numToChar(namedEndCol + numCols);

          name.text = this.joinRange({
            start: this.joinRef(namedStart),
            end: this.joinRef(namedEnd),
          });
        }
      } else {
        const namedRef = this.splitRef(ref);
        const namedCol = _charToNum(namedRef.col);

        if (namedRef.row === currentRow && currentCol < namedCol) {
          namedRef.col = _numToChar(namedCol + numCols);

          name.text = this.joinRef(namedRef);
        }
      }
    });
  }

  // Look for any merged cell, named table or named range definitions below
  // `currentRow` and push down by `numRows` (used when rows are inserted).
  pushDown(workbook, sheet, tables, currentRow, numRows) {
    const mergeCells = sheet.find("mergeCells");

    // Update merged cells below this row
    sheet.findall("mergeCells/mergeCell").forEach((mergeCell) => {
      const mergeRange = this.splitRange(mergeCell.attrib.ref);
      const mergeStart = this.splitRef(mergeRange.start);
      const mergeEnd = this.splitRef(mergeRange.end);

      if (mergeStart.row > currentRow) {
        mergeStart.row += numRows;
        mergeEnd.row += numRows;

        mergeCell.attrib.ref = this.joinRange({
          start: this.joinRef(mergeStart),
          end: this.joinRef(mergeEnd),
        });
      }

      //add new merge cell
      if (mergeStart.row == currentRow) {
        for (let i = 1; i <= numRows; i++) {
          const newMergeCell = this.cloneElement(mergeCell);
          mergeStart.row += 1;
          mergeEnd.row += 1;
          newMergeCell.attrib.ref = this.joinRange({
            start: this.joinRef(mergeStart),
            end: this.joinRef(mergeEnd),
          });
          mergeCells.attrib.count += 1;
          mergeCells._children.push(newMergeCell);
        }
      }
    });

    // Update named tables below this row
    tables.forEach((table) => {
      const tableRoot = table.root;
      const tableRange = this.splitRange(tableRoot.attrib.ref);
      const tableStart = this.splitRef(tableRange.start);
      const tableEnd = this.splitRef(tableRange.end);

      if (tableStart.row > currentRow) {
        tableStart.row += numRows;
        tableEnd.row += numRows;

        tableRoot.attrib.ref = this.joinRange({
          start: this.joinRef(tableStart),
          end: this.joinRef(tableEnd),
        });

        const autoFilter = tableRoot.find("autoFilter");
        if (autoFilter !== null) {
          // XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = tableRoot.attrib.ref;
        }
      }
    });

    // Named cells/ranges
    workbook.findall("definedNames/definedName").forEach((name) => {
      const ref = name.text;
      if (this.isRange(ref)) {
        const namedRange = this.splitRange(ref); //TODO : I think is there a bug, the ref is equal to [sheetName]![startRange]:[endRange]
        const namedStart = this.splitRef(namedRange.start); // here, namedRange.start is [sheetName]![startRange] ?
        const namedEnd = this.splitRef(namedRange.end);
        if (namedStart) {
          if (namedStart.row > currentRow) {
            namedStart.row += numRows;
            namedEnd.row += numRows;

            name.text = this.joinRange({
              start: this.joinRef(namedStart),
              end: this.joinRef(namedEnd),
            });
          }
        }
        if (this.option && this.option.pushDownPageBreakOnTableSubstitution) {
          if (
            this.sheet.name == name.text.split("!")[0].replace(/'/gi, "") &&
            namedEnd
          ) {
            if (namedEnd.row > currentRow) {
              namedEnd.row += numRows;
              name.text = this.joinRange({
                start: this.joinRef(namedStart),
                end: this.joinRef(namedEnd),
              });
            }
          }
        }
      } else {
        const namedRef = this.splitRef(ref);
        if (namedRef.row > currentRow) {
          namedRef.row += numRows;
          name.text = this.joinRef(namedRef);
        }
      }
    });
  }

  getWidthCell(numCol, sheet) {
    const defaultWidth =
      sheet.root.find("sheetFormatPr").attrib["defaultColWidth"];
    if (!defaultWidth) {
      // TODO : Check why defaultColWidth is not set ?
      defaultWidth = 11.42578125;
    }
    let finalWidth = defaultWidth;
    sheet.root.findall("cols/col").forEach((col) => {
      if (numCol >= col.attrib["min"] && numCol <= col.attrib["max"]) {
        if (col.attrib["width"] != undefined) {
          finalWidth = col.attrib["width"];
        }
      }
    });
    return Number.parseFloat(finalWidth);
  }

  getWidthMergeCell(mergeCell, sheet) {
    let mergeWidth = 0;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartCol = _charToNum(this.splitRef(mergeRange.start).col);
    const mergeEndCol = _charToNum(this.splitRef(mergeRange.end).col);
    for (let i = mergeStartCol; i < mergeEndCol + 1; i++) {
      mergeWidth += this.getWidthCell(i, sheet);
    }
    return mergeWidth;
  }

  getHeightCell(numRow, sheet) {
    let finalHeight =
      sheet.root.find("sheetFormatPr").attrib["defaultRowHeight"];
    sheet.root.findall("sheetData/row").forEach((row) => {
      if (numRow == row.attrib["r"]) {
        if (row.attrib["ht"] != undefined) {
          finalHeight = row.attrib["ht"];
        }
      }
    });
    return Number.parseFloat(finalHeight);
  }

  getHeightMergeCell(mergeCell, sheet) {
    let mergeHeight = 0;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    for (let i = mergeStartRow; i < mergeEndRow + 1; i++) {
      mergeHeight += this.getHeightCell(i, sheet);
    }
    return mergeHeight;
  }

  getNbRowOfMergeCell(mergeCell) {
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    return mergeEndRow - mergeStartRow + 1;
  }

  pixelsToEMUs(pixels) {
    return Math.round((pixels * 914400) / 96);
  }

  columnWidthToEMUs(width) {
    // TODO : This is not the true. Change with true calcul
    // can find help here :
    // https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
    // https://stackoverflow.com/questions/58021996/how-to-set-the-fixed-column-width-values-in-inches-apache-poi
    // https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Sheet.html#setColumnWidth-int-int-
    // https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
    // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    // http://lcorneliussen.de/raw/dashboards/ooxml/
    return this.pixelsToEMUs(width * 7.625579987895905);
  }

  rowHeightToEMUs(height) {
    // TODO : need to be verify
    return Math.round((height / 72) * 914400);
  }

  findMaxFileId(fileNameRegex, idRegex) {
    const files = this.archive.file(fileNameRegex);
    const maxFile = files.reduce((p, c) => {
      if (p == null) {
        return c.name;
      }
      return p.name > c.name ? p.name : c.name;
    }, null);
    let maxid = 0;
    if (maxFile != null) {
      maxid = idRegex.exec(maxFile)[1];
    }
    maxid++;
    return maxid;
  }

  cellInMergeCells(cell, mergeCell) {
    const cellCol = _charToNum(this.splitRef(cell.attrib.r).col);
    const cellRow = this.splitRef(cell.attrib.r).row;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartCol = _charToNum(this.splitRef(mergeRange.start).col);
    const mergeEndCol = _charToNum(this.splitRef(mergeRange.end).col);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    if (cellCol >= mergeStartCol && cellCol <= mergeEndCol) {
      if (cellRow >= mergeStartRow && cellRow <= mergeEndRow) {
        return true;
      }
    }
    return false;
  }

  imageToBuffer(imageObj) {
    if (!imageObj) {
      throw new TypeError("imageObj cannot be null");
    }
    if (imageObj instanceof Buffer) {
      return _checkImage(imageObj);
    }
    if (typeof imageObj === "string" || imageObj instanceof String) {
      try {
        imageObj = imageObj.toString();
        const imagePath = this.option.imageRootPath
          ? `${this.option.imageRootPath}/${imageObj}`
          : imageObj;
        if (fs.existsSync(imagePath)) {
          return _checkImage(
            Buffer.from(
              fs.readFileSync(imagePath, { encoding: "base64" }),
              "base64"
            )
          );
        }
        return _checkImage(Buffer.from(imageObj, "base64"));
      } catch (error) {
        throw new TypeError("imageObj cannot be parse as a buffer");
      }
    }
    throw new TypeError(`imageObj type is not supported : ${typeof imageObj}`);
  }

  findMaxId(element, tag, attr, idRegex) {
    let maxId = 0;
    element.findall(tag).forEach((element) => {
      const match = idRegex.exec(element.attrib[attr]);
      if (match == null) {
        throw new Error("Can not find the id!");
      }
      const cid = parseInt(match[1]);
      if (cid > maxId) {
        maxId = cid;
      }
    });
    return ++maxId;
  }
}

module.exports = Workbook;
