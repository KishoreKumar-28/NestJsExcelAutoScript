import {
  HttpException,
  HttpStatus,
  Inject,
  Injectable,
  UploadedFile,
} from '@nestjs/common';
import { Pool } from 'pg';
import { constants } from './include';
import * as ExcelJS from 'exceljs';
@Injectable()
export class AppService {
  constructor(@Inject('PG_POOL') private readonly pool: Pool) {}
  async uploadFile(@UploadedFile() file: Express.Multer.File) {
    try {
      // Intially disable all foreign key constraints in tables
      await this.pool.query(constants.disableForeignKeyQuery);
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(file.buffer);
      const sheets = workbook.worksheets;
      const pageIdToNameMap: { [pageName: string]: string } = {};
      const dropdownSourceKeyValuePairs: { [key: string]: any } = {};
      //--Col Id to Col name mapping
      const colData = {}; // Object to store Col ID, Page Type, Page ID, Col Name and Col Dropdown source
      //Process All Languages sheet to find the Row Id of English language
      const allLanguagesSheet = sheets.find(
        (sheet) => sheet.name === constants.allLanguages,
      );
      // Dynamically find the header row and store the language and row header index of all languages sheet.
      let languageheaderRowIndex = constants.index;
      let languageheaderColIndex = constants.index;
      let rowHeaderAllLanguageColIndex = constants.index;
      let englishRowId;
      const languageHeader = await this.findHeaderRowAndColIndex(
        allLanguagesSheet,
        constants.language,
      );
      languageheaderRowIndex = languageHeader.headerRowIndex;
      languageheaderColIndex = languageHeader.headerColIndex;
      rowHeaderAllLanguageColIndex = (
        await this.findHeaderRowAndColIndex(allLanguagesSheet, constants.rowId)
      ).headerColIndex;
      if (
        languageheaderRowIndex !== constants.index &&
        languageheaderColIndex !== constants.index
      ) {
        // Find the English in language column and fetch the Row Id against the value.
        for (
          let rowIndx = languageheaderRowIndex + constants.one;
          rowIndx <= allLanguagesSheet.lastRow.number;
          rowIndx++
        ) {
          const rowCell = allLanguagesSheet
            .getCell(rowIndx, languageheaderColIndex)
            .value.toString();
          if (rowCell == constants.english) {
            englishRowId = allLanguagesSheet
              .getCell(rowIndx, rowHeaderAllLanguageColIndex)
              .value.toString();
          }
        }
      }
      // Find the 'All Cols' sheet
      const allColsSheet = sheets.find(
        (sheet) => sheet.name === constants.allCols,
      );
      // Find 'All Tokens' sheet
      const allTokensSheet = sheets.find(
        (sheet) => sheet.name === constants.allTokens,
      );
      // Find 'All Labels' sheet
      const allLabelsSheet = sheets.find(
        (sheet) => sheet.name === constants.allLabels,
      );
      // Find 'All Units' sheet
      const allUnitsSheet = sheets.find(
        (sheet) => sheet.name === constants.allUnits,
      );
      // Process the "All Cols" sheet to store the Col ID, Page Type, Page ID, Col Name, COl Data type, Col Dropdown source, Col Status.
      if (!allColsSheet) {
        throw new Error(constants.allColsError);
      }
      let colIdIndex = constants.index;
      let pageTypeIndex = constants.index;
      let pageIdIndex = constants.index;
      let colNameIndex = constants.index;
      let colDataTypeIndex = constants.index;
      let colDropDownSourceIndex = constants.index;
      let colStatusIndex = constants.index;
      let headerRowIndex = constants.index;
      // Find the indices of the headers in "All Cols" sheet.
      for (
        let rowIndex = constants.one;
        rowIndex <= allColsSheet.lastRow.number;
        rowIndex++
      ) {
        const row = allColsSheet.getRow(rowIndex);
        for (
          let colIndex = constants.one;
          colIndex <= row.cellCount;
          colIndex++
        ) {
          const cellValue = row.getCell(colIndex).value?.toString();
          if (cellValue && constants.colId.test(cellValue)) {
            colIdIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.pageType.test(cellValue)) {
            pageTypeIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.pageId.test(cellValue)) {
            pageIdIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.colName.test(cellValue)) {
            colNameIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.colDataType.test(cellValue)) {
            colDataTypeIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.colDropDownSource.test(cellValue)) {
            colDropDownSourceIndex = colIndex;
            headerRowIndex = rowIndex;
          }
          if (cellValue && constants.colStatus === cellValue) {
            colStatusIndex = colIndex;
            headerRowIndex = rowIndex;
          }
        }
        // Exit the loop once the header is found
        if (headerRowIndex !== constants.index) break;
      }
      // If any of the index is not found then throw header error.
      if (
        colIdIndex === constants.index ||
        pageTypeIndex === constants.index ||
        pageIdIndex === constants.index ||
        colNameIndex === constants.index ||
        colDataTypeIndex === constants.index ||
        colDropDownSourceIndex === constants.index ||
        colStatusIndex === constants.index
      ) {
        throw new Error(constants.headerError);
      }
      // Read the data under the headers and store it in the object
      for (
        let rowIndex = headerRowIndex + constants.one;
        rowIndex <= allColsSheet.lastRow.number;
        rowIndex++
      ) {
        const row = allColsSheet.getRow(rowIndex);
        const colId = row.getCell(colIdIndex).value?.toString();
        const pageType = row.getCell(pageTypeIndex).value?.toString();
        const pageId = row.getCell(pageIdIndex).value?.toString();
        const colName = row.getCell(colNameIndex).value?.toString();
        const colDataType = row.getCell(colDataTypeIndex).value?.toString();
        const colDropDownSource = row
          .getCell(colDropDownSourceIndex)
          .value?.toString();
        const colStatus = row.getCell(colStatusIndex).value?.toString();
        // Store all the values into an object of Col ID as a key.
        if (colId && colName) {
          colData[colId] = {
            pageType,
            pageId,
            colName,
            colDataType,
            colDropDownSource,
            colStatus,
          };
        }
      }
      // Function to store the key value pair of datatype and row ID in all tokens
      const dataTypeToRowId = {};
      const objectTypeToRowId = {};
      const statusesToRowId = {};
      let rowAllTokensColIndex = constants.index;
      let rowTypeAllTokensColIndex = constants.index;
      // Dynamically find the header row
      let allTokensheaderRowIndex = constants.index;
      allTokensheaderRowIndex = (
        await this.findHeaderRowAndColIndex(
          allTokensSheet,
          constants.tokenPattern,
        )
      ).headerRowIndex;
      // If header Row Index is not found then throw header error.
      if (allTokensheaderRowIndex === constants.index) {
        console.log(constants.headerError + allTokensSheet.name);
      }
      // Find the start and end column index for the merged "Token" header along with Url Type Row ID, DDS type Row ID and Formula Type Row ID.
      let tokenColStartIndex = constants.index;
      let tokenColEndIndex = constants.index;
      let urlTypeRowID = null;
      let ddsTypeRowID = null;
      let formulaTypeRowID = null;
      if (allTokensheaderRowIndex !== constants.index) {
        const headerRow = allTokensSheet.getRow(allTokensheaderRowIndex);
        for (
          let sheetColIndex = constants.one;
          sheetColIndex <= allTokensSheet.lastColumn.number;
          sheetColIndex++
        ) {
          const cell = headerRow.getCell(sheetColIndex);
          if (
            cell.value &&
            constants.tokenPattern.test(cell.value.toString())
          ) {
            if (tokenColStartIndex === constants.index) {
              tokenColStartIndex = sheetColIndex;
            }
            tokenColEndIndex = sheetColIndex;
          }
          if (cell.value && constants.rowId.test(cell.value.toString())) {
            rowAllTokensColIndex = sheetColIndex;
          }
          if (cell.value && constants.rowType.test(cell.value.toString())) {
            rowTypeAllTokensColIndex = sheetColIndex;
          }
        }
        if (
          tokenColStartIndex === constants.index ||
          tokenColEndIndex === constants.index ||
          rowAllTokensColIndex === constants.index
        ) {
          console.log(constants.allTokenIndexError + allTokensSheet.name);
        }
      }
      //Dynamically find the header row of all labels sheet.
      let allLabelsHeaderRowIndex = constants.index;
      for (
        let rowIndex = constants.one;
        rowIndex <= allLabelsSheet.lastRow.number;
        rowIndex++
      ) {
        const row = allLabelsSheet.getRow(rowIndex);
        for (
          let colIndex = constants.one;
          colIndex <= row.cellCount;
          colIndex++
        ) {
          const cellValue = row.getCell(colIndex).value?.toString();
          if (cellValue && constants.label.test(cellValue)) {
            allLabelsHeaderRowIndex = rowIndex;
            break;
          }
        }
      }
      // Find the start and end columns for the merged "Label" header in all labels sheet
      let labelColStartIndex = constants.index;
      let labelColEndIndex = constants.index;
      let rowAllLabelsColIndex = constants.index;
      let valueDefaultDataColIndex = constants.index;
      if (allLabelsHeaderRowIndex !== constants.index) {
        const headerRow = allLabelsSheet.getRow(allLabelsHeaderRowIndex);
        for (
          let sheetColIndex = constants.one;
          sheetColIndex <= allLabelsSheet.lastColumn.number;
          sheetColIndex++
        ) {
          const cell = headerRow.getCell(sheetColIndex);
          if (cell.value && constants.label.test(cell.value.toString())) {
            if (labelColStartIndex === constants.index) {
              labelColStartIndex = sheetColIndex;
            }
            labelColEndIndex = sheetColIndex;
          }
          if (cell.value && constants.rowId.test(cell.value.toString())) {
            rowAllLabelsColIndex = sheetColIndex;
          }
          if (
            cell.value &&
            constants.titemColumns.valueDefaultData === cell.value.toString()
          ) {
            valueDefaultDataColIndex = sheetColIndex;
          }
        }
        // If any of index is not found then throw error.
        if (
          labelColStartIndex === constants.index ||
          labelColEndIndex === constants.index ||
          rowAllLabelsColIndex === constants.index ||
          valueDefaultDataColIndex === constants.index
        ) {
          console.log(constants.allLabelsIndexError + allLabelsSheet.name);
        }
      }
      //Find and Store the URL, Validate Data, Formula Type & ddsType Row ID
      for (
        let i = allTokensheaderRowIndex + constants.one;
        i <= allTokensSheet.lastRow.number;
        i++
      ) {
        const row = allTokensSheet.getRow(i);
        for (let j = tokenColStartIndex; j <= tokenColEndIndex; j++) {
          const cell = row.getCell(j);
          if (cell.value && cell.value.toString() === constants.urltype) {
            urlTypeRowID = row.getCell(rowAllTokensColIndex);
          }
          if (cell.value && cell.value.toString() === constants.ddstype) {
            ddsTypeRowID = row.getCell(rowAllTokensColIndex).value.toString();
          }
          if (cell.value && cell.value.toString() === constants.validateData) {
            formulaTypeRowID = row
              .getCell(rowAllTokensColIndex)
              .value.toString();
          }
          if (
            urlTypeRowID !== null &&
            ddsTypeRowID !== null &&
            formulaTypeRowID !== null
          ) {
            break;
          }
        }
      }
      // Identify the "DataType", "UserType", "Statuses" and "Object" row dynamically within the "Token" header columns
      let dataTypeRowIndex = constants.index;
      let dataTypeColIndex = constants.index;
      let userTypeRowIndex = constants.index;
      let objectTypeRowIndex = constants.index;
      let objectTypeColIndex = constants.index;
      let statuesRowIndex = constants.index;
      let statuesColIndex = constants.index;
      if (
        tokenColStartIndex !== constants.index &&
        tokenColEndIndex !== constants.index
      ) {
        for (
          let i = allTokensheaderRowIndex + constants.one;
          i <= allTokensSheet.lastRow.number;
          i++
        ) {
          const row = allTokensSheet.getRow(i);
          for (let j = tokenColStartIndex; j <= tokenColEndIndex; j++) {
            const cell = row.getCell(j);
            if (cell.value && constants.userType === cell.value.toString()) {
              userTypeRowIndex = i;
            }
            if (cell.value && constants.objectType === cell.value.toString()) {
              objectTypeRowIndex = i;
              objectTypeColIndex = j;
            }
            if (cell.value && constants.dataType.test(cell.value.toString())) {
              dataTypeRowIndex = i;
              dataTypeColIndex = j;
            }
            if (cell.value && constants.statuses === cell.value.toString()) {
              statuesRowIndex = i;
              statuesColIndex = j;
            }
          }
          // If all the index is found then break the loop.
          if (
            dataTypeRowIndex !== constants.index &&
            dataTypeColIndex !== constants.index &&
            userTypeRowIndex !== constants.index &&
            objectTypeRowIndex !== constants.index &&
            objectTypeColIndex !== constants.index &&
            statuesRowIndex !== constants.index &&
            statuesColIndex !== constants.index
          )
            break;
        }
        // If any of the index is not found then throw error.
        if (
          dataTypeRowIndex === constants.index ||
          userTypeRowIndex === constants.index ||
          objectTypeRowIndex === constants.index ||
          objectTypeColIndex === constants.index ||
          statuesRowIndex === constants.index ||
          statuesColIndex === constants.index
        ) {
          console.log(constants.datatypeError + allTokensSheet.name);
        }
      }
      // Collect values from the "Token" header under the "DataType" row.
      if (
        dataTypeRowIndex !== constants.index &&
        tokenColStartIndex !== constants.index &&
        tokenColEndIndex !== constants.index &&
        rowAllTokensColIndex !== constants.index
      ) {
        const dataTypeResult = await this.findValuesAndRowIdInAllTokens(
          allTokensSheet,
          dataTypeRowIndex,
          dataTypeColIndex,
          tokenColEndIndex,
          rowAllTokensColIndex,
        );
        // Create key-value pairs with RowType as key and RowId as value.
        for (let i = constants.zero; i < dataTypeResult.value.length; i++) {
          dataTypeToRowId[dataTypeResult.value[i]] =
            dataTypeResult.rowIdOfValue[i];
        }
      }
      // Find the values under Object to store as key value pair for tFormat.ObjectType.
      if (
        objectTypeRowIndex !== constants.index &&
        objectTypeColIndex !== constants.index &&
        tokenColStartIndex !== constants.index &&
        tokenColEndIndex !== constants.index &&
        rowAllTokensColIndex !== constants.index
      ) {
        const objectTypeResult = await this.findValuesAndRowIdInAllTokens(
          allTokensSheet,
          objectTypeRowIndex,
          objectTypeColIndex,
          tokenColEndIndex,
          rowAllTokensColIndex,
        );
        // Create key-value pairs with Object as key and RowId as value.
        for (let i = 0; i < objectTypeResult.value.length; i++) {
          objectTypeToRowId[objectTypeResult.value[i]] =
            objectTypeResult.rowIdOfValue[i];
        }
      }
      // Find the values under Statuses to store as key value pair for tFormat.Status.
      if (
        statuesRowIndex !== constants.index &&
        statuesColIndex !== constants.index &&
        tokenColStartIndex !== constants.index &&
        tokenColEndIndex !== constants.index &&
        rowAllTokensColIndex !== constants.index
      ) {
        const statusesResult = await this.findValuesAndRowIdInAllTokens(
          allTokensSheet,
          statuesRowIndex,
          statuesColIndex,
          tokenColEndIndex,
          rowAllTokensColIndex,
        );
        // Create key-value pairs with Status as key and RowId as value.
        for (let i = 0; i < statusesResult.value.length; i++) {
          statusesToRowId[statusesResult.value[i]] =
            statusesResult.rowIdOfValue[i];
        }
      }
      //Find the index of Admin and default page expand level in all labels sheet with the default data.
      let adminRowIndex = constants.index;
      let defaultExpandLevelRowIndex = constants.index;
      let foundAdmin = false;
      let foundDefaultExpand = false;
      for (let j = labelColStartIndex; j <= labelColEndIndex; j++) {
        for (
          let i = allLabelsHeaderRowIndex + constants.one;
          i <= allLabelsSheet.lastRow.number;
          i++
        ) {
          const row = allLabelsSheet.getRow(i);
          const cell = row.getCell(j);
          if (cell.value != null && cell.value != undefined) {
            if (cell.value.toString() === constants.admin && !foundAdmin) {
              adminRowIndex = i;
              foundAdmin = true;
            } else if (
              cell.value.toString() === constants.defaultExpandLevel &&
              !foundDefaultExpand
            ) {
              defaultExpandLevelRowIndex = i;
              foundDefaultExpand = true;
            }
          }
          if (foundAdmin && foundDefaultExpand) break; // Break the inner loop
        }
        if (foundAdmin && foundDefaultExpand) break; // Break the outer loop
      }
      // Find the default data by fetching the cell in Admin Row and value Default data column
      const userId = allLabelsSheet
        .getCell(adminRowIndex, valueDefaultDataColIndex)
        .value.toString();
      const defaultPgExpandLevel = allLabelsSheet
        .getCell(defaultExpandLevelRowIndex, valueDefaultDataColIndex)
        .value.toString();
      let userTypeRowId = null;
      let adminUser;
      // Check all tokens sheet for the Default User type
      for (let i = userTypeRowIndex; i <= allTokensSheet.lastRow.number; i++) {
        for (let j = tokenColStartIndex; j <= tokenColEndIndex; j++) {
          const row = allTokensSheet.getRow(i);
          const cell = row.getCell(j);
          if (cell.value != null && cell.value != undefined) {
            if (cell.value.toString() === constants.adminUserType) {
              userTypeRowId = row
                .getCell(rowAllTokensColIndex)
                .value.toString();
              break;
            }
          }
        }
        if (userTypeRowId !== null) break;
      }
      //Insert the Amin User into tUser table
      if (
        userTypeRowId !== null &&
        userTypeRowId != undefined &&
        userId !== null &&
        userId != undefined
      ) {
        const inserttUserQuery = {
          text: constants.inserttUserQuery,
          values: [userId, userTypeRowId],
        };
        const adminUserRecord = await this.pool.query(inserttUserQuery);
        adminUser = adminUserRecord.rows[0].User;
      } else {
        console.log(constants.userNotFoundError);
      }
      // Insert the tRow.Row with zero for default data cell creation.
      const insertDefaultTRowQuery = {
        text: constants.insertDefaulttRowQuery,
        values: [constants.zero, constants.one],
      };
      await this.pool.query(insertDefaultTRowQuery);
      // Process through all the sheets in the Excel file
      for (const sheet of sheets) {
        // Process only the mentioned sheets in include file
        if (constants.sheetNames.includes(sheet.name)) {
          console.log(constants.process + sheet.name);
          // Process 'All Pages' sheet
          if (sheet.name === constants.allPages) {
            // Initialize arrays to store page IDs and names
            const pageIds: string[] = [];
            const pageNames: string[] = [];
            // Iterate through each cell in the sheet
            for (
              let sheetRowIndex = constants.one;
              sheetRowIndex <= sheet.lastRow.number;
              sheetRowIndex++
            ) {
              for (
                let sheetColIndex = constants.one;
                sheetColIndex <= sheet.lastColumn.number;
                sheetColIndex++
              ) {
                const cell = sheet.getCell(sheetRowIndex, sheetColIndex);
                // Check for page ID pattern and populate pageIds array
                if (
                  cell.value &&
                  constants.pageIdMandatory.test(cell.value.toString())
                ) {
                  for (
                    let rowIdx = constants.one;
                    rowIdx <= sheet.lastRow.number;
                    rowIdx++
                  ) {
                    const rowCell = sheet.getCell(rowIdx, sheetColIndex);
                    const value = rowCell.value;
                    if (value !== null && value !== undefined) {
                      pageIds.push(value.toString());
                    }
                  }
                }
                // Check for page name pattern and populate pageNames array
                if (
                  cell.value &&
                  constants.pageName.test(cell.value.toString())
                ) {
                  for (
                    let rowIdx = constants.one;
                    rowIdx <= sheet.lastRow.number;
                    rowIdx++
                  ) {
                    const rowCell = sheet.getCell(rowIdx, sheetColIndex);
                    const value = rowCell.value;
                    if (value !== null && value !== undefined) {
                      pageNames.push(value.toString());
                    }
                  }
                }
              }
            }
            // Create a key-value pair of page ID and page name
            for (let i = constants.zero; i < pageIds.length; i++) {
              pageIdToNameMap[pageNames[i]] = pageIds[i];
            }
          }
          // Check if sheet name is a key in pageIdToNameMap
          if (sheet.name in pageIdToNameMap) {
            var pageId = pageIdToNameMap[sheet.name];
          }
          // Find header row index based on specific constants
          let headerRowIndex = constants.index;
          for (let i = constants.one; i <= sheet.lastRow.number; i++) {
            const row = sheet.getRow(i);
            for (let j = constants.one; j <= row.cellCount; j++) {
              const cell = row.getCell(j);
              if (cell.value && constants.rowType.test(cell.value.toString())) {
                headerRowIndex = i;
                break;
              }
            }
            // If Header Row Index is found then break the loop.
            if (headerRowIndex !== constants.index) break;
          }
          // Log error if headerRowIndex is still index constant.
          if (headerRowIndex === constants.index) {
            console.log(constants.headerError + sheet.name);
            continue; // Skip to the next sheet
          }
          // Retrieve header row using header row index.
          const headerRow = sheet.getRow(headerRowIndex);
          // Initialize variables for column indices and nested column.
          let rowIdColumnIndex = constants.index;
          let rowStatusColumnIndex = constants.index;
          let nestedColumnStartIndex = constants.index;
          let nestedColumnEndIndex = constants.index;
          let pageIdColumnIndex = constants.index;
          let colIdColumnIndex = constants.index;
          let pageTypeIndex = constants.index;
          let colStatusColIndex = constants.index;
          let colFormulaColIndex = constants.index;
          let colCommentColIndex = constants.index;
          let colOwnerColIndex = constants.index;
          let nestedColumn = constants.nestedColumns[sheet.name];
          // Iterate through header row to identify specific columns
          for (
            let sheetColIndex = constants.one;
            sheetColIndex <= sheet.lastColumn.number;
            sheetColIndex++
          ) {
            const cell = headerRow.getCell(sheetColIndex);
            const cellValue = cell.value ? cell.value.toString().trim() : null;
            if (cellValue && cellValue !== null) {
              if (constants.rowId.test(cellValue)) {
                rowIdColumnIndex = sheetColIndex;
              } else if (constants.rowStatus.test(cellValue)) {
                rowStatusColumnIndex = sheetColIndex;
              } else if (constants.pageId.test(cellValue)) {
                pageIdColumnIndex = sheetColIndex;
              } else if (constants.colId.test(cellValue)) {
                colIdColumnIndex = sheetColIndex;
              } else if (constants.pageType.test(cellValue)) {
                pageTypeIndex = sheetColIndex;
              } else if (constants.colComment === cellValue) {
                colCommentColIndex = sheetColIndex;
              } else if (constants.colStatus === cellValue) {
                colStatusColIndex = sheetColIndex;
              } else if (constants.colFormula === cellValue) {
                colFormulaColIndex = sheetColIndex;
              } else if (constants.colOwner === cellValue) {
                colOwnerColIndex = sheetColIndex;
              } else if (
                nestedColumn &&
                new RegExp(nestedColumn).test(cellValue)
              ) {
                if (nestedColumnStartIndex === constants.index) {
                  nestedColumnStartIndex = sheetColIndex;
                }
                nestedColumnEndIndex = sheetColIndex;
              }
            }
          }
          // Log error if rowStatusColumnIndex is still index constant
          if (rowStatusColumnIndex === constants.index) {
            console.log(constants.rowStatusError + sheet.name);
            continue; // Skip to the next sheet
          }
          // Initialize arrays and objects to store previous rows and last row at level
          let previousRows = [];
          let lastRowAtLevel = {};
          const sharedColumnQueries = [];
          let colOrder = constants.one;
          let processedPageId = null;
          let startColOrderAfterEachPage;
          // Iterate through each row in the sheet
          for (
            let rowIdx = headerRowIndex + constants.one;
            rowIdx <= sheet.lastRow.number;
            rowIdx++
          ) {
            const row = sheet.getRow(rowIdx);
            // Check if the row is empty
            let isRowEmpty = true;
            for (
              let colIdx = constants.one;
              colIdx <= row.cellCount;
              colIdx++
            ) {
              const cell = row.getCell(colIdx);
              if (
                cell.value !== null &&
                cell.value !== undefined &&
                cell.value.toString().trim() !== ''
              ) {
                isRowEmpty = false;
                break;
              }
            }
            // Skip the empty rows
            if (isRowEmpty) {
              continue;
            }
            // Retrieve Row ID, Row Status, Page ID, Col ID, Page Type, Col Status, Col Comment, Col Formula, Col Owner values
            const rowIdCell =
              rowIdColumnIndex !== constants.index
                ? row.getCell(rowIdColumnIndex)
                : null;
            let rowValue = rowIdCell ? rowIdCell.value : null;
            const rowStatusValue = this.getCellValue(row, rowStatusColumnIndex);
            const pageIdValue = this.getCellValue(row, pageIdColumnIndex);
            const colIdValue = this.getCellValue(row, colIdColumnIndex);
            const pageTypeValue = this.getCellValue(row, pageTypeIndex);
            const colStatusValue = this.getCellValue(row, colStatusColIndex);
            const colCommentValue = this.getCellValue(row, colCommentColIndex);
            const colFormulaValue = this.getCellValue(row, colFormulaColIndex);
            const colOwnerValue = this.getCellValue(row, colOwnerColIndex);
            // Check the Page ID is present and Sheet name is "All Pages" then insert the record into tPg table followed by tFormat table.
            let insertedtFormatIdForPage;
            if (
              pageIdValue !== null &&
              pageIdValue !== undefined &&
              sheet.name === constants.allPages
            ) {
              const inserttPgQuery = {
                text: constants.inserttPgQuery,
                values: [pageIdValue],
              };
              try {
                await this.pool.query(inserttPgQuery);
              } catch (error) {
                console.error(constants.tPgError, error);
              }
              // Method to Find the Nested Column ID in a page.
              let nestedColId;
              for (const colId in colData) {
                const col = colData[colId];
                if (
                  col.pageId === pageIdValue &&
                  col.colStatus.includes(constants.nested)
                ) {
                  nestedColId = colId;
                  break;
                }
              }
              //Insert the tFormat record for the PageID inserted into tPg table
              const pgObjectType = objectTypeToRowId[constants.page];
              const inserttFormatForPageQuery = {
                text: constants.inserttFormatForPageQuery,
                values: [
                  adminUser,
                  pgObjectType,
                  pageIdValue,
                  defaultPgExpandLevel,
                  nestedColId,
                ],
              };
              try {
                const pagetFormatRecord = await this.pool.query(
                  inserttFormatForPageQuery,
                );
                insertedtFormatIdForPage = pagetFormatRecord.rows[0].Format;
              } catch (error) {
                console.error(constants.tFormatForPageError, error);
              }
            }
            // Check the ColId value is present and sheet name is "all Cols" then insert record into tCol table followed by tFormat table.
            if (
              colIdValue !== null &&
              colIdValue !== undefined &&
              sheet.name === constants.allCols
            ) {
              const insertTColQuery = {
                text: constants.inserttColQuery,
                values: [colIdValue],
              };
              try {
                await this.pool.query(insertTColQuery);
              } catch (error) {
                console.error(constants.tPgError, error);
              }
              const statusIds = [];
              // If the Col owner is admin then store the Row ID of admin in ColOwner.
              const colOwner =
                colOwnerValue === constants.admin ? adminUser : null;
              //Find the Col Status Token IDs by splitting the Items in the cell with semicolon.
              const cellValues = colStatusValue
                .split(constants.semicolon)
                .map((val) => val.trim())
                .filter(Boolean);
              if (cellValues.length > 0) {
                for (const value of cellValues) {
                  statusIds.push(statusesToRowId[value]);
                }
              }
              // Convert the formula into json using the formula type as a key and formula as pair.
              const colFormula =
                colFormulaValue !== null
                  ? JSON.stringify({
                      [formulaTypeRowID]: colFormulaValue,
                    })
                  : null;
              // Column comment in json with english Row ID
              const colCommentJson =
                colCommentValue !== null
                  ? JSON.stringify({
                      [englishRowId]: colCommentValue,
                    })
                  : null;
              let insertedtFormatIdForCol;
              // Store the shared column queries to insert the records for each pages in the excel sheet.
              const colObjectType = objectTypeToRowId[constants.column];
              if (pageTypeValue === constants.eachPage) {
                const inserttFormatForColQuery = {
                  text: constants.inserttFormatForColQuery,
                  values: [
                    adminUser,
                    colObjectType,
                    colIdValue,
                    colOrder,
                    colOwner,
                    statusIds,
                    colFormula,
                    colCommentJson,
                  ],
                };
                sharedColumnQueries.push(inserttFormatForColQuery);
                colOrder++;
              }
              // If the Page Id is found then start inserting the shared columns with the container column.
              if (pageIdValue !== null && pageIdValue !== undefined) {
                // If the new page Id is found on the Page ID column then insert the shared column queries.
                if (
                  pageIdValue !== processedPageId ||
                  processedPageId === null
                ) {
                  const updateQuery = constants.updateAnyColumnsIntFormatQuery(
                    constants.container,
                  );
                  for (const query of sharedColumnQueries) {
                    try {
                      const columntFormatRecord = await this.pool.query(query);
                      insertedtFormatIdForCol =
                        columntFormatRecord.rows[0].Format;
                      const updatetFormatColumnQuery = {
                        text: updateQuery,
                        values: [pageIdValue, insertedtFormatIdForCol],
                      };
                      await this.pool.query(updatetFormatColumnQuery);
                    } catch (error) {
                      console.error(constants.tFormatForColumnError, error);
                    }
                  }
                  // After inserting the shared columns, insert the Page columns
                  const inserttFormatForColQuery = {
                    text: constants.inserttFormatForColQuery,
                    values: [
                      adminUser,
                      colObjectType,
                      colIdValue,
                      colOrder,
                      colOwner,
                      statusIds,
                      colFormula,
                      colCommentJson,
                    ],
                  };
                  const columntFormatRecord = await this.pool.query(
                    inserttFormatForColQuery,
                  );
                  insertedtFormatIdForCol = columntFormatRecord.rows[0].Format;
                  processedPageId = pageIdValue;
                  startColOrderAfterEachPage = colOrder + constants.one;
                }
                //If the Page Id is different from the last processed page Id then insert the Page columns without shared columns.
                else {
                  const inserttFormatForColQuery = {
                    text: constants.inserttFormatForColQuery,
                    values: [
                      adminUser,
                      colObjectType,
                      colIdValue,
                      startColOrderAfterEachPage,
                      colOwner,
                      statusIds,
                      colFormula,
                      colCommentJson,
                    ],
                  };
                  const columntFormatRecord = await this.pool.query(
                    inserttFormatForColQuery,
                  );
                  insertedtFormatIdForCol = columntFormatRecord.rows[0].Format;
                  processedPageId = pageIdValue;
                  startColOrderAfterEachPage++;
                }
              }
            }
            // if Row value is N/A then skip the row for tRow insertion.
            if (
              rowValue !== null &&
              rowValue !== undefined &&
              rowIdColumnIndex !== constants.index
            ) {
              if (rowValue.toString() === constants.nonInsertRow) continue;
            }
            // If the Row column is exist and no Row value then generate the Row value.
            if (
              rowIdColumnIndex !== constants.index &&
              (rowValue === null || rowValue === undefined)
            ) {
              const nextRowValue = await this.getNextRowValue();
              rowValue = nextRowValue;
            }
            // Determine row level based on row status and nested columns
            let rowLevel = constants.one;
            // If the Row status column contains section head then Row level is 0.
            if (
              rowStatusValue !== null &&
              rowStatusValue !== undefined &&
              rowStatusValue.toString() === constants.sectionHead
            ) {
              rowLevel = constants.zero;
            }
            // Else count the cells in the nested column present in the iterating sheet.
            else if (
              nestedColumnStartIndex !== constants.index &&
              nestedColumnEndIndex !== constants.index
            ) {
              for (
                let colIdx = nestedColumnStartIndex;
                colIdx <= nestedColumnEndIndex;
                colIdx++
              ) {
                const cell = row.getCell(colIdx);
                if (cell.value) {
                  rowLevel = colIdx - nestedColumnStartIndex + constants.one;
                  break;
                }
              }
            }
            // Initialize parent and sibling row IDs
            let parentRowId = null;
            let siblingRowId = null;
            // Determine parent and sibling row IDs based on previous rows
            for (
              let i = previousRows.length - constants.one;
              i >= constants.zero;
              i--
            ) {
              if (previousRows[i].rowLevel < rowLevel) {
                parentRowId = previousRows[i].id;
                break;
              }
            }
            // Store the Parent Row Id with the Row Level.
            const lastRowKey = `${parentRowId}-${rowLevel}`;
            if (lastRowAtLevel[lastRowKey]) {
              siblingRowId = lastRowAtLevel[lastRowKey].id;
            }
            let newRowId = null;
            let savedRowEntity;
            // Insert the row based on row value or generate a new row value if there is no Row column.
            if (rowValue !== null && rowValue !== undefined) {
              const inserttRowQuery = {
                text: constants.inserttRowQuery,
                values: [
                  Number(rowValue),
                  Number(pageId),
                  rowLevel,
                  parentRowId,
                ],
              };
              try {
                // Execute the insert query and save the new row ID in savedRowEntity.
                const result = await this.pool.query(inserttRowQuery);
                savedRowEntity = result.rows[0].Row;
              } catch (error) {
                console.error(constants.rowError, error);
                throw error;
              }
            } else {
              const nextRowValue = await this.getNextRowValue();
              const inserttRowQuery = {
                text: constants.inserttRowQuery,
                values: [
                  Number(nextRowValue),
                  Number(pageId),
                  rowLevel,
                  parentRowId,
                ],
              };
              try {
                // Execute the insert query and save the new row ID in savedRowEntity.
                const result = await this.pool.query(inserttRowQuery);
                savedRowEntity = result.rows[0].Row;
              } catch (error) {
                console.error(constants.rowError, error);
                throw error;
              }
            }
            try {
              // Save the new row entity in tRow and retrieve the new row ID
              newRowId = savedRowEntity;
              // Store the Row value with the Token as a key value pair to validate with the DropDown Source column
              if (rowLevel === constants.zero) {
                dropdownSourceKeyValuePairs[
                  row.getCell(nestedColumnStartIndex).value.toString()
                ] = savedRowEntity;
                // console.log(dropdownSourceKeyValuePairs);
              }
              // Log error when newRowId is undefined
              if (newRowId === undefined) {
                console.error(constants.emptyRowError);
                continue;
              }
              // Update the sibling Row ID for inserted Row.
              const updateSiblingRowIntRowQuery = {
                text: constants.updateSiblingRowIntRowQuery,
                values: [newRowId, siblingRowId],
              };
              try {
                await this.pool.query(updateSiblingRowIntRowQuery);
              } catch (error) {
                console.error(constants.siblingRowUpdateError, error);
                throw error;
              }
              // Store current row details in previousRows and lastRowAtLevel objects
              previousRows.push({
                id: newRowId,
                rowValue,
                rowLevel,
                parentRowId,
                siblingRowId,
              });
              lastRowAtLevel[lastRowKey] = {
                id: newRowId,
                rowValue,
                rowLevel,
                parentRowId,
              };
            } catch (err) {
              console.error(constants.rowError + err);
              continue; // Skip to the next row in case of error
            }
            //Insert tFormat record for every row inserted into tRow table
            let insertedtFormatIdForRow;
            const rowObjectType =
              objectTypeToRowId[constants.tformatColumns.rowId];
            const inserttFormatForRowQuery = {
              text: constants.inserttFormatForRowQuery,
              values: [adminUser, rowObjectType, savedRowEntity, adminUser],
            };
            try {
              const insertedtFormatRecord = await this.pool.query(
                inserttFormatForRowQuery,
              );
              insertedtFormatIdForRow = insertedtFormatRecord.rows[0].Format;
            } catch (error) {
              console.error(constants.tFormatForRowError, error);
              throw error;
            }
            // Check every cell in inserted row that is present in tItemColumns or tFormatColumns to insert into tCell, tItem and tFormat
            for (
              let colIdx = constants.one;
              colIdx <= row.cellCount;
              colIdx++
            ) {
              let isTitemColumn = false;
              let isTformatColumn = false;
              let colID;
              let colDataType;
              let colDropDownSource;
              let savedCellEntity;
              const cellObjectType = objectTypeToRowId[constants.cell];
              const cell: any = sheet.getCell(rowIdx, colIdx).value;
              let cellValue: any = null;
              if (cell != null && cell != undefined) {
                if (typeof cell === constants.object) {
                  // Handle different object types
                  if (constants.richText in cell) {
                    // If the cell contains rich text, concatenate all the text parts
                    cellValue = cell.richText
                      .map((part: any) => part.text)
                      .join('');
                  }
                } else {
                  // If the cell value is a simple type, use it directly
                  cellValue = cell.toString();
                }
                if (cellValue instanceof Date) {
                  // If the cell value is a Date, format it to 'YYYY-MM-DD'
                  cellValue = cellValue.toISOString().split(constants.t)[0];
                }
                // Get the header cell value for the current column index
                const headerCell = sheet.getRow(headerRowIndex).getCell(colIdx);
                let headerCellValue = headerCell.value?.toString().trim();
                // Remove trailing '*' if present
                if (headerCellValue?.endsWith(constants.star)) {
                  headerCellValue = headerCellValue.slice(
                    constants.zero,
                    constants.index,
                  ); // Remove the last character
                }
                // Check the tItemColumns in include file contains the header Cell value.
                if (headerCellValue != null || headerCellValue != undefined) {
                  for (const key in constants.titemColumns) {
                    if (constants.titemColumns[key] === headerCellValue) {
                      isTitemColumn = true;
                      break;
                    }
                  }
                  // Check the tFormatColumns in include file contains the header Cell value.
                  for (const key in constants.tformatColumns) {
                    if (constants.tformatColumns[key] === headerCellValue) {
                      isTformatColumn = true;
                      break;
                    }
                  }
                  if (isTitemColumn) {
                    // Validate the headers with the Col name in "All Cols" sheet and the page Id matches or Page type is each page.
                    for (const key in colData) {
                      if (
                        colData[key].colName === headerCellValue &&
                        (colData[key].pageId == pageId ||
                          colData[key].pageType == constants.eachPage)
                      ) {
                        colID = key;
                        colDataType = colData[key].colDataType;
                        colDropDownSource = colData[key].colDropDownSource;
                        break;
                      }
                    }
                    // For Default column insert tCell record with tRow.Row = 0
                    let rowEntityValue;
                    if (
                      constants.titemColumns.colDefaultData ===
                        headerCellValue ||
                      constants.titemColumns.valueDefaultData ===
                        headerCellValue
                    ) {
                      rowEntityValue = constants.zero;
                    } else {
                      rowEntityValue = savedRowEntity;
                    }
                    // Insert the Col ID and Row ID into tCell table
                    if (colID != null && savedRowEntity != null) {
                      const inserttCellQuery = {
                        text: constants.inserttCellQuery,
                        values: [colID, rowEntityValue],
                      };
                      try {
                        // Execute the insert query and return the saved cell entity
                        const result = await this.pool.query(inserttCellQuery);
                        savedCellEntity = result.rows[0];
                      } catch (error) {
                        console.error(constants.tCellRecordError, error);
                        throw error;
                      }
                      if (
                        savedCellEntity != null &&
                        savedCellEntity != undefined
                      ) {
                        // Object type of Item to store in tFormat.ObjectType
                        const itemObjectType =
                          objectTypeToRowId[constants.item];
                        if (
                          constants.titemColumns.pageType === headerCellValue ||
                          constants.titemColumns.rowType === headerCellValue ||
                          constants.titemColumns.pageEdition ===
                            headerCellValue ||
                          constants.titemColumns.colDataType ===
                            headerCellValue ||
                          constants.titemColumns.valueDataType ===
                            headerCellValue ||
                          constants.titemColumns.valueStatus === headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          // Using the dropdownSource present in the page type column and cell value find the Token ID of the cell value
                          if (colDropDownSource && dataType) {
                            // Step 1: Find the row in allTokensSheet where RowType is "node" and the Token value is colDropDownSource
                            let tokenRowIndex = constants.index;
                            for (
                              let rowIndex = constants.one;
                              rowIndex <= allTokensSheet.lastRow.number;
                              rowIndex++
                            ) {
                              const row = allTokensSheet.getRow(rowIndex);
                              const rowTypeCell = row.getCell(
                                rowTypeAllTokensColIndex,
                              );
                              const tokenCell = row.getCell(tokenColStartIndex);
                              if (
                                rowTypeCell &&
                                tokenCell &&
                                constants.node.test(
                                  rowTypeCell.value?.toString(),
                                ) &&
                                tokenCell.value?.toString() ===
                                  colDropDownSource
                              ) {
                                tokenRowIndex = rowIndex;
                                break;
                              }
                            }
                            if (tokenRowIndex !== constants.index) {
                              // Step 2: Retrieve the values under the found token and store them with their Row IDs
                              const tokenValueToRowIdMap: Record<
                                string,
                                string
                              > = {};
                              let toBreak = false;
                              for (
                                let rowIdx = tokenRowIndex + constants.one;
                                rowIdx <= allTokensSheet.lastRow.number;
                                rowIdx++
                              ) {
                                for (
                                  let colIndex = tokenColStartIndex;
                                  colIndex <= tokenColEndIndex;
                                  colIndex++
                                ) {
                                  const cell = allTokensSheet
                                    .getRow(rowIdx)
                                    .getCell(colIndex);
                                  if (
                                    colIndex === tokenColStartIndex &&
                                    cell.value != null &&
                                    cell.value != undefined
                                  ) {
                                    toBreak = true;
                                    break;
                                  }
                                  if (toBreak) break;
                                  if (cell && cell.value) {
                                    const tokenValue = cell.value.toString();
                                    const rowIDCell = allTokensSheet
                                      .getRow(rowIdx)
                                      .getCell(rowAllTokensColIndex);
                                    if (rowIDCell && rowIDCell.value) {
                                      const rowIDValue =
                                        rowIDCell.value.toString();
                                      tokenValueToRowIdMap[tokenValue] =
                                        rowIDValue;
                                    }
                                  }
                                }
                              }
                              // Step 3: Split cellValue and map each value to corresponding RowID
                              const cellValues = cellValue.split(
                                constants.semicolon,
                              );
                              const matchedRowIds = cellValues
                                .map((val) => tokenValueToRowIdMap[val.trim()])
                                .filter(Boolean);
                              if (matchedRowIds.length > constants.zero) {
                                try {
                                  // Insert corresponding tItem records and collect Item IDs
                                  const itemIds = [];
                                  let itemOrder = constants.one;
                                  for (const rowIDValue of matchedRowIds) {
                                    const inserttItemWithObjectQuery = {
                                      text: constants.inserttItemWithObjectQuery,
                                      values: [dataType, Number(rowIDValue)],
                                    };
                                    try {
                                      // Execute the insert query and return the new row ID
                                      const result = await this.pool.query(
                                        inserttItemWithObjectQuery,
                                      );
                                      const insertedItemId =
                                        result.rows[0].Item;
                                      if (
                                        insertedItemId !== null &&
                                        matchedRowIds.length >= constants.two
                                      ) {
                                        const inserttFormatForItemQuery = {
                                          text: constants.inserttFormatForItemQuery,
                                          values: [
                                            adminUser,
                                            itemObjectType,
                                            insertedItemId,
                                            itemOrder,
                                            adminUser,
                                          ],
                                        };
                                        await this.pool.query(
                                          inserttFormatForItemQuery,
                                        );
                                      }
                                      itemIds.push(insertedItemId);
                                      itemOrder++;
                                    } catch (error) {
                                      console.error(
                                        constants.itemIdError,
                                        error,
                                      );
                                      throw error;
                                    }
                                  }
                                  // Update the saved cell entity with the array of Item IDs
                                  await this.updateItemIdsIntCell(
                                    savedCellEntity.Cell,
                                    itemIds,
                                  );
                                } catch (error) {
                                  console.error(constants.itemIdError, error);
                                }
                              }
                            }
                          }
                        }
                        if (
                          constants.titemColumns.pageId === headerCellValue ||
                          constants.titemColumns.colId === headerCellValue
                        ) {
                          // Find the Row Id of datatype from dataTypeToRowId
                          const dataType = dataTypeToRowId[colDataType];
                          try {
                            const itemIds = [];
                            const inserttItemWithObjectQuery = {
                              text: constants.inserttItemWithObjectQuery,
                              values: [dataType, Number(cellValue)],
                            };
                            try {
                              // Execute the insert query and return the new row ID
                              const result = await this.pool.query(
                                inserttItemWithObjectQuery,
                              );
                              const insertedItemId = result.rows[0].Item;
                              itemIds.push(insertedItemId);
                            } catch (error) {
                              console.error(constants.rowError, error);
                              throw error;
                            }
                            // Update the saved cell entity with the array of Item IDs
                            await this.updateItemIdsIntCell(
                              savedCellEntity.Cell,
                              itemIds,
                            );
                          } catch (error) {
                            console.error(constants.itemIdError, error);
                          }
                        }
                        // If the header cell value matches the Release date then insert the value in DateTime column in tItem table
                        if (
                          constants.titemColumns.releaseDate === headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          try {
                            const itemIds = [];
                            const inserttItemWithDateTimeQuery = {
                              text: constants.inserttItemWithDateTimeQuery,
                              values: [dataType, cellValue],
                            };
                            try {
                              // Execute the insert query and return the new row ID
                              const result = await this.pool.query(
                                inserttItemWithDateTimeQuery,
                              );
                              const insertedItemId = result.rows[0].Item;
                              itemIds.push(insertedItemId);
                            } catch (error) {
                              console.error(constants.rowError, error);
                              throw error;
                            }
                            // Update the saved cell entity with the array of Item IDs
                            await this.updateItemIdsIntCell(
                              savedCellEntity.Cell,
                              itemIds,
                            );
                          } catch (error) {
                            console.error(constants.itemIdError, error);
                          }
                        }
                        // If the headerCell value is matches the Unit factor column name then insert the value into number column in tItem table.
                        if (
                          constants.titemColumns.unitFactor === headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          try {
                            // Insert corresponding tItem records and collect Item IDs
                            const itemIds = [];
                            const inserttItemWithNumberQuery = {
                              text: constants.inserttItemWithNumberQuery,
                              values: [dataType, Number(cellValue)],
                            };
                            try {
                              // Execute the insert query and return the new row ID
                              const result = await this.pool.query(
                                inserttItemWithNumberQuery,
                              );
                              const insertedItemId = result.rows[0].Item;
                              itemIds.push(insertedItemId);
                            } catch (error) {
                              console.error(constants.rowError, error);
                              throw error;
                            }
                            // Update the saved cell entity with the array of Item IDs
                            await this.updateItemIdsIntCell(
                              savedCellEntity.Cell,
                              itemIds,
                            );
                          } catch (error) {
                            console.error(constants.itemIdError, error);
                          }
                        }
                        // If headerCell value matches the Col DropDown source or Value Dropdown Source column then insert the Dropdown source Token ID from DDS pages.
                        if (
                          constants.titemColumns.colDropDownSource ===
                            headerCellValue ||
                          constants.titemColumns.valueDropdownSource ===
                            headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          let dropdownSource = null;
                          const cellValues = cellValue
                            .split(constants.semicolon)
                            .map((val) => val.trim())
                            .filter(Boolean);
                          if (cellValues.length > 0) {
                            try {
                              // Insert corresponding tItem records and collect Item IDs.
                              const itemIds = [];
                              for (const value of cellValues) {
                                dropdownSource = null;
                                // Check DropDown source is avaliable in All Tokens sheet
                                for (
                                  let j = tokenColStartIndex;
                                  j <= tokenColEndIndex;
                                  j++
                                ) {
                                  for (
                                    let i = headerRowIndex + constants.one;
                                    i <= allTokensSheet.lastRow.number;
                                    i++
                                  ) {
                                    const row = allTokensSheet.getRow(i);
                                    const cell = row.getCell(j);
                                    if (
                                      cell.value &&
                                      cell.value.toString() === value
                                    ) {
                                      dropdownSource = row
                                        .getCell(rowAllTokensColIndex)
                                        .value.toString();
                                      break;
                                    }
                                  }
                                  if (dropdownSource !== null) break;
                                }
                                // If not found in All Tokens sheet, check all Labels sheet
                                if (dropdownSource === null) {
                                  for (
                                    let j = labelColStartIndex;
                                    j <= labelColEndIndex;
                                    j++
                                  ) {
                                    for (
                                      let i =
                                        allLabelsHeaderRowIndex + constants.one;
                                      i <= allLabelsSheet.lastRow.number;
                                      i++
                                    ) {
                                      const row = allLabelsSheet.getRow(i);
                                      const cell = row.getCell(j);
                                      if (
                                        cell.value &&
                                        cell.value.toString() === value
                                      ) {
                                        dropdownSource = row
                                          .getCell(rowAllLabelsColIndex)
                                          .value.toString();
                                        break;
                                      }
                                    }
                                    if (dropdownSource !== null) break;
                                  }
                                }
                                // If not found in all labels sheet, check all Unit sheet
                                if (dropdownSource === null) {
                                  const unitHeaderIndex =
                                    await this.findHeaderRowAndColIndex(
                                      allUnitsSheet,
                                      constants.unit,
                                    );
                                  const rowHeaderIndex =
                                    await this.findHeaderRowAndColIndex(
                                      allUnitsSheet,
                                      constants.rowId,
                                    );
                                  for (
                                    let i =
                                      unitHeaderIndex.headerRowIndex +
                                      constants.one;
                                    i <= allUnitsSheet.lastRow.number;
                                    i++
                                  ) {
                                    const unitCellValue = allUnitsSheet.getCell(
                                      i,
                                      unitHeaderIndex.headerColIndex,
                                    );
                                    if (
                                      unitCellValue.value &&
                                      unitCellValue.value.toString() === value
                                    ) {
                                      dropdownSource = allUnitsSheet
                                        .getCell(
                                          i,
                                          rowHeaderIndex.headerColIndex,
                                        )
                                        .value.toString();
                                    }
                                  }
                                }
                                // If not found then check all the section-Head Rows key value pair
                                if (dropdownSource === null) {
                                  dropdownSource =
                                    dropdownSourceKeyValuePairs[value];
                                }
                                // Process found dropdownSource
                                if (
                                  dataType !== null &&
                                  dropdownSource !== null &&
                                  dropdownSource !== undefined &&
                                  ddsTypeRowID !== null
                                ) {
                                  const json = JSON.stringify({
                                    [ddsTypeRowID]: dropdownSource,
                                  });
                                  const insertedItemId =
                                    await this.insertItemWithJson(
                                      dataType,
                                      json,
                                    );
                                  itemIds.push(insertedItemId);
                                }
                              }
                              // Update the saved tCell record with the array of Item IDs
                              if (itemIds.length > 0) {
                                await this.updateItemIdsIntCell(
                                  savedCellEntity.Cell,
                                  itemIds,
                                );
                              } else {
                                console.error(
                                  constants.noItemIdtoUpdatetCellError,
                                );
                              }
                            } catch (error) {
                              console.error(constants.itemIdError, error);
                            }
                          }
                        }
                        if (
                          constants.titemColumns.pageName === headerCellValue ||
                          constants.titemColumns.pageSEO === headerCellValue ||
                          constants.titemColumns.colName === headerCellValue ||
                          constants.titemColumns.language === headerCellValue ||
                          constants.titemColumns.region === headerCellValue ||
                          constants.titemColumns.supplier === headerCellValue ||
                          constants.titemColumns.token === headerCellValue ||
                          constants.titemColumns.model === headerCellValue ||
                          constants.titemColumns.unit === headerCellValue ||
                          constants.titemColumns.labels === headerCellValue ||
                          constants.titemColumns.valueFormula ===
                            headerCellValue ||
                          constants.titemColumns.colDefaultData ===
                            headerCellValue ||
                          constants.titemColumns.valueDefaultData ===
                            headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          // Split cellValue and insert each split value into tItem
                          const cellValues = cellValue
                            .split(constants.semicolon)
                            .map((val) => val.trim())
                            .filter(Boolean);
                          if (cellValues.length > 0) {
                            try {
                              // Insert corresponding tItem records and collect Item IDs
                              const itemIds = [];
                              let itemOrder = constants.one;
                              for (const value of cellValues) {
                                const json = JSON.stringify({
                                  [englishRowId]: value,
                                });
                                const insertedItemId =
                                  await this.insertItemWithJson(dataType, json);
                                if (
                                  insertedItemId !== null &&
                                  cellValues.length >= constants.two
                                ) {
                                  const inserttFormatForItemQuery = {
                                    text: constants.inserttFormatForItemQuery,
                                    values: [
                                      adminUser,
                                      itemObjectType,
                                      insertedItemId,
                                      itemOrder,
                                      adminUser,
                                    ],
                                  };
                                  await this.pool.query(
                                    inserttFormatForItemQuery,
                                  );
                                }
                                itemIds.push(insertedItemId);
                                itemOrder++;
                              }
                              // Update the saved cell entity with the array of Item IDs
                              await this.updateItemIdsIntCell(
                                savedCellEntity.Cell,
                                itemIds,
                              );
                            } catch (error) {
                              console.error(constants.itemIdError, error);
                            }
                          }
                        }
                        if (
                          constants.titemColumns.pageURL === headerCellValue
                        ) {
                          const dataType = dataTypeToRowId[colDataType];
                          try {
                            const itemIds = [];
                            // Insert corresponding tItem records and collect Item IDs
                            const json = JSON.stringify({
                              [urlTypeRowID]: cellValue,
                            });
                            const insertedItemId =
                              await this.insertItemWithJson(dataType, json);
                            itemIds.push(insertedItemId);
                            // Update the saved cell entity with the array of Item IDs
                            await this.updateItemIdsIntCell(
                              savedCellEntity.Cell,
                              itemIds,
                            );
                          } catch (error) {
                            console.error(constants.itemIdError, error);
                          }
                        }
                      }
                    }
                  }
                  if (isTformatColumn) {
                    if (
                      constants.tformatColumns.rowComment === headerCellValue &&
                      insertedtFormatIdForRow !== null &&
                      insertedtFormatIdForRow !== undefined
                    ) {
                      const updateQuery =
                        constants.updateAnyColumnsIntFormatQuery(
                          constants.comment,
                        );
                      const json = JSON.stringify({
                        [englishRowId]: cellValue,
                      });
                      // Update the tFormat with the Row commnet on tFormat with Comment column
                      const updatetFormatColumnQuery = {
                        text: updateQuery,
                        values: [json, insertedtFormatIdForRow],
                      };
                      try {
                        await this.pool.query(updatetFormatColumnQuery);
                      } catch (error) {
                        console.error(
                          constants.tFormatUpdateError,
                          constants.comment,
                          error,
                        );
                      }
                    }
                    if (
                      constants.tformatColumns.rowStatus === headerCellValue &&
                      insertedtFormatIdForRow !== null &&
                      insertedtFormatIdForRow !== undefined
                    ) {
                      const cellValues = cellValue
                        .split(constants.semicolon)
                        .map((val) => val.trim())
                        .filter(Boolean);
                      if (cellValues.length > 0) {
                        const statusIds = [];
                        for (const value of cellValues) {
                          statusIds.push(statusesToRowId[value]);
                        }
                        const updateQuery =
                          constants.updateAnyColumnsIntFormatQuery(
                            constants.status,
                          );
                        // Update the tFormat with the Row Status on tFormat with Status column
                        const updatetFormatColumnQuery = {
                          text: updateQuery,
                          values: [statusIds, insertedtFormatIdForRow],
                        };
                        try {
                          await this.pool.query(updatetFormatColumnQuery);
                        } catch (error) {
                          console.error(
                            constants.tFormatUpdateError,
                            constants.status,
                            error,
                          );
                        }
                      }
                    }
                    if (
                      constants.tformatColumns.pageOwner === headerCellValue &&
                      insertedtFormatIdForPage !== null &&
                      insertedtFormatIdForPage !== undefined
                    ) {
                      if (cellValue === constants.admin) {
                        const updateQuery =
                          constants.updateAnyColumnsIntFormatQuery(
                            constants.owner,
                          );
                        // Update the tFormat with the page owner on tFormat with Owner column
                        const updatetFormatColumnQuery = {
                          text: updateQuery,
                          values: [adminUser, insertedtFormatIdForPage],
                        };
                        try {
                          await this.pool.query(updatetFormatColumnQuery);
                        } catch (error) {
                          console.error(
                            constants.tFormatUpdateError,
                            constants.status,
                            error,
                          );
                        }
                      }
                    }
                    if (
                      constants.tformatColumns.pageComment ===
                        headerCellValue &&
                      insertedtFormatIdForPage !== null &&
                      insertedtFormatIdForPage !== undefined
                    ) {
                      const updateQuery =
                        constants.updateAnyColumnsIntFormatQuery(
                          constants.comment,
                        );
                      const json = JSON.stringify({
                        [englishRowId]: cellValue,
                      });
                      // Update the tFormat with the Page commnet on tFormat with Comment column
                      const updatetFormatColumnQuery = {
                        text: updateQuery,
                        values: [json, insertedtFormatIdForPage],
                      };
                      try {
                        await this.pool.query(updatetFormatColumnQuery);
                      } catch (error) {
                        console.error(
                          constants.tFormatUpdateError,
                          constants.comment,
                          error,
                        );
                      }
                    }
                    if (
                      constants.tformatColumns.pageStatus === headerCellValue &&
                      insertedtFormatIdForPage !== null &&
                      insertedtFormatIdForPage !== undefined
                    ) {
                      const cellValues = cellValue
                        .split(constants.semicolon)
                        .map((val) => val.trim())
                        .filter(Boolean);
                      if (cellValues.length > 0) {
                        const statusIds = [];
                        for (const value of cellValues) {
                          statusIds.push(statusesToRowId[value]);
                        }
                        const updateQuery =
                          constants.updateAnyColumnsIntFormatQuery(
                            constants.status,
                          );
                        // Construct the query object for updating tFormat with Comment column
                        const updatetFormatColumnQuery = {
                          text: updateQuery,
                          values: [statusIds, insertedtFormatIdForPage],
                        };
                        try {
                          await this.pool.query(updatetFormatColumnQuery);
                        } catch (error) {
                          console.error(
                            constants.tFormatUpdateError,
                            constants.status,
                            error,
                          );
                        }
                      }
                    }
                    // Insert the Cell Id of the Default Cell into tFormat.Default when the header cell value matches the Value Data type or Col Data type.
                    if (
                      (constants.tformatColumns.valueDefaultData ===
                        headerCellValue ||
                        constants.tformatColumns.colDefaultData ===
                          headerCellValue) &&
                      savedCellEntity !== null
                    ) {
                      const inserttFormatForDefaultColQuery = {
                        text: constants.inserttFormatForDefaultColQuery,
                        values: [
                          adminUser,
                          cellObjectType,
                          savedCellEntity.Cell,
                          savedCellEntity.Cell,
                          adminUser,
                        ],
                      };
                      try {
                        await this.pool.query(inserttFormatForDefaultColQuery);
                      } catch (error) {
                        console.error(
                          constants.tFormatUpdateError,
                          constants.status,
                          error,
                        );
                      }
                    }
                  }
                }
              }
            }
          }
          // Update sibling rows to null for lastChildRow
          for (let key in lastRowAtLevel) {
            const rowId = lastRowAtLevel[key].id;
            const updateSiblingRowIntRowToNull = {
              text: constants.updateSiblingRowIntRowToNull,
              values: [rowId],
            };
            try {
              // Execute the update query
              await this.pool.query(updateSiblingRowIntRowToNull);
            } catch (error) {
              console.error(constants.updatingSiblingRowToNullError, error);
              throw error;
            }
          }
        }
      }
      await this.pool.query(constants.enableForeignKeyQuery);
      return { message: constants.successMessage };
    } catch (error) {
      // log the error and throw HTTP exception
      console.error(error);
      throw new HttpException(
        constants.serverError,
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }
  // Method to Generate the next Row value from the last Row.
  async getNextRowValue() {
    const getLastRowIdQuery = { text: constants.getLastRowId };
    try {
      const result = await this.pool.query(getLastRowIdQuery);
      const lastRow = result.rows[0];
      return lastRow ? parseInt(lastRow.Row) + constants.one : constants.one;
    } catch (error) {
      console.error(constants.lastRowFetchingError, error);
      throw error;
    }
  }
  // Method to Update the ItemIds into tCell table.
  async updateItemIdsIntCell(cellId: number, newItemIds: number[]) {
    const gettCellQuery = {
      text: constants.gettCellQuery,
      values: [cellId],
    };
    try {
      const result = await this.pool.query(gettCellQuery);
      const currentItems = result.rows[0]?.Items || [];
      // Append the new ItemIDs to the array
      const updatedItems = [...currentItems, ...newItemIds];
      const updateItemIdsIntCellQuery = {
        text: constants.updateItemIdsIntCellQuery,
        values: [updatedItems, cellId],
      };
      // Execute the update query
      await this.pool.query(updateItemIdsIntCellQuery);
    } catch (error) {
      console.error(constants.tCellUpdateError, error);
      throw error;
    }
  }
  // Method to Insert record into tItem table with JSON column.
  async insertItemWithJson(dataType: number, json: string) {
    const insertItemWithJsonQuery = {
      text: constants.inserttItemWithJsonQuery,
      values: [dataType, json],
    };
    try {
      // Execute the insert query and return the new row ID
      const result = await this.pool.query(insertItemWithJsonQuery);
      return result.rows[0].Item;
    } catch (error) {
      console.error(constants.rowError, error);
      throw error;
    }
  }
  // Method to find the Token ID and Value for any DDS header.
  async findValuesAndRowIdInAllTokens(
    sheet: any,
    headerRowIndex: number,
    headerColIndex: number,
    tokenColEndIndex: number,
    rowAllTokensColIndex: number,
  ) {
    let shouldBreak = false;
    const value = [];
    const rowIdOfValue = [];
    for (
      let i = headerRowIndex + constants.one;
      i <= sheet.lastRow.number;
      i++
    ) {
      const row = sheet.getRow(i);
      for (let j = headerColIndex; j <= tokenColEndIndex; j++) {
        const cell = row.getCell(j);
        let tokenValue;
        if (cell.value != null || cell.value != undefined) {
          tokenValue = cell.value.toString();
        }
        // Check if there's a corresponding RowId value in the same row
        const rowCell = row.getCell(rowAllTokensColIndex);
        const rowValue = rowCell ? rowCell.value : null;
        // Break if a value is found in the same column index as "DataType"
        if (
          j === headerColIndex &&
          cell.value != null &&
          cell.value != undefined
        ) {
          shouldBreak = true;
          break;
        }
        if (shouldBreak) break;
        // Store the hierarchy and row value
        if (
          rowValue !== null &&
          rowValue !== undefined &&
          tokenValue !== null &&
          tokenValue !== undefined
        ) {
          value.push(tokenValue);
          rowIdOfValue.push(rowValue.toString());
        }
      }
    }
    return { value, rowIdOfValue };
  }
  // Method to find the Header Row and column index with input of sheet name and header name.
  async findHeaderRowAndColIndex(sheet: any, header: any) {
    let headerRowIndex;
    let headerColIndex;
    for (let i = constants.one; i <= sheet.lastRow.number; i++) {
      const row = sheet.getRow(i);
      for (let j = constants.one; j <= row.cellCount; j++) {
        const cell = row.getCell(j);
        if (cell.value && header.test(cell.value.toString())) {
          headerRowIndex = i;
          headerColIndex = j;
          break;
        }
      }
    }
    return { headerRowIndex, headerColIndex };
  }
  // Method to fetch the value from the row using the row and Column Index.
  private getCellValue(row: any, columnIndex: number): string | null {
    if (columnIndex === constants.index) {
      return null;
    }
    let cellValue = row.getCell(columnIndex).value;
    if (cellValue !== null && cellValue !== undefined) {
      if (typeof cellValue === constants.object) {
        if (constants.richText in cellValue) {
          cellValue = cellValue.richText.map((part: any) => part.text).join('');
        }
      } else {
        cellValue = cellValue.toString();
      }
      return cellValue;
    }
    return null;
  }
}
