import { Maybe } from "@/entities/Maybe";
import { deepCopyObj } from "@/functions/deepCopyObj";
import Log from "@/modules/Log/Log";
import { unifyCode } from "@/widgetsBusiness/ExcelSupplierCodesCheck";
export module ExcelDataFilter {
  export type HeadersInterface = {
    name: string;
    index: number;
    columnRawValues: any[];
    declinedColumnValues: any[];
  };
  export type HeadersMap = Map<HeadersInterface["name"], HeadersInterface>;
  export type ExcelColumnValuesByHeaderGetterArg = {
    selectedColumnHeader: HeadersInterface;
    isViewOnlyInSelectedRange: boolean;
    headersRowIndex: number;
    selectedRangeRowIndex: number;
  };

  export type getFilteredExcelDataArg = {
    chosenHeaders: HeadersInterface[];
    headersArr: HeadersInterface[];
  };

  export type deleteDeclinedRowsByHeadersArg = {
    clearedValuesByColumns: any[][];
    headersArr: HeadersInterface[];
    headersArrNames: any[];
    chosenHeaders: HeadersInterface[];
  };

  export type getDeclinedHeadersIndexesArg = {
    chosenHeadersNames: string[];
    headersRow: any[];
  };

  export type HeadersCheckArg = {
    clearedValuesByColumns: any[][];
    headersArrNames: string[];
  };

  export type getFilteredRowsByHeaderColumnValuesArg = {
    values: any[][];
    headerRowNamesToReplaceHeaderRow: Maybe<string | number | boolean>[];
    declinedHeadersColumnValuesArr: HeadersInterface[];
    skipRowsQuantity: number;
  };
}

export class ExcelDataFilter {
  async getFilteredExcelData(
    args: ExcelDataFilter.getFilteredExcelDataArg
  ): Promise<any[]> {
    try {
      const { chosenHeaders, headersArr } = args;
      const context = await Excel.run<Excel.RequestContext>(
        async (context: Excel.RequestContext) => context
      );
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;

      const headersArrNames = headersArr.map((elem) => {
        return elem.name;
      });
      const clearedValuesByColumns = this.getFilteredRowsByHeaderColumnValues({
        values,
        headerRowNamesToReplaceHeaderRow: headersArrNames,
        declinedHeadersColumnValuesArr: headersArr,
      });
      const clearedValuesByHeaders = this.deleteDeclinedRowsByHeaders({
        clearedValuesByColumns,
        headersArr,
        headersArrNames,
        chosenHeaders,
      });

      return clearedValuesByHeaders;
    } catch (e) {
      throw Log.error("/loadColumnValues", e);
    }
  }

  deleteDeclinedRowsByHeaders(
    args: ExcelDataFilter.deleteDeclinedRowsByHeadersArg
  ): any[][] {
    const {
      clearedValuesByColumns,
      headersArr,
      headersArrNames,
      chosenHeaders,
    } = args;

    const chosenHeadersNames = chosenHeaders.map((elem) => {
      return elem.name;
    });

    const checkedValues = this.headersCheck({
      clearedValuesByColumns,
      headersArrNames,
    });

    const clearedValuesByHeaders: any[][] = deepCopyObj(checkedValues);
    const headersRow = checkedValues[0];

    const declinedHeadersIndexes = this.getDeclinedHeadersIndexes({
      chosenHeadersNames,
      headersRow,
    });

    for (let m = 0; m < declinedHeadersIndexes.length; m++) {
      const declinedHeadersIndex: Maybe<any> = declinedHeadersIndexes[m];
      if (declinedHeadersIndex != null) {
        for (let i = 0; i < clearedValuesByHeaders.length; i++) {
          clearedValuesByHeaders[i].splice(declinedHeadersIndexes[m], 1);
        }
      }
    }
    return clearedValuesByHeaders;
  }

  getDeclinedHeadersIndexes(
    args: ExcelDataFilter.getDeclinedHeadersIndexesArg
  ): number[] {
    const { chosenHeadersNames, headersRow } = args;
    const declinedHeadersIndexesRow: number[] = [];

    for (let i = 0; i < headersRow.length; i++) {
      const header: Maybe<any> = headersRow[i];
      if (header != null) {
        if (!chosenHeadersNames.includes(header)) {
          declinedHeadersIndexesRow.push(i);
        }
      }
    }
    const declinedHeadersIndexes = declinedHeadersIndexesRow.map(
      (elem, index) => {
        return elem - +index;
      }
    );
    return declinedHeadersIndexes;
  }

  headersCheck(args: ExcelDataFilter.HeadersCheckArg): any[][] {
    const { clearedValuesByColumns, headersArrNames } = args;
    const headersRow: Maybe<any[]> = clearedValuesByColumns[0];
    const headersNames = deepCopyObj(headersArrNames);

    if (headersRow.length > headersNames.length) {
      const difference = headersRow.length - headersNames.length;
      // headersNames.fill('', headersNames.length)
      for (let i = 0; i < difference; i++) {
        headersNames.push("");
      }
    }

    const checkedClearedValuesByColumns = deepCopyObj(clearedValuesByColumns);
    const isheadersRowEqualToHeadersNames =
      unifyCode(JSON.stringify(headersRow)) !=
      unifyCode(JSON.stringify(headersNames));
    if (isheadersRowEqualToHeadersNames)
      checkedClearedValuesByColumns[0] = headersNames;
    return checkedClearedValuesByColumns;
  }

  getFilteredRowsByHeaderColumnValues(
    args: ExcelDataFilter.getFilteredRowsByHeaderColumnValuesArg
  ): any[][] {
    const {
      values,
      declinedHeadersColumnValuesArr,
      headerRowNamesToReplaceHeaderRow,
      skipRowsQuantity,
    } = args;
    const clearedValuesByColumns = deepCopyObj(values);

    for (let index = 0; index < skipRowsQuantity; index++) {
      clearedValuesByColumns.shift();
    }

    for (const headerObj of declinedHeadersColumnValuesArr) {
      for (
        let rowIndex = 0;
        rowIndex < clearedValuesByColumns.length;
        rowIndex++
      ) {
        const declinedValues = headerObj.declinedColumnValues;

        if (declinedValues.length == 0) continue;

        const removeRow = () => {
          clearedValuesByColumns.splice(rowIndex, 1);
          rowIndex--;
        };

        const rawRowValues: Maybe<any[]> = clearedValuesByColumns[rowIndex];

        if (rawRowValues == null || !rawRowValues.length) {
          removeRow();
          continue;
        }

        const rowValue: Maybe<any> = rawRowValues[headerObj.index];
        if (rowValue == null) {
          removeRow();
          continue;
        }

        const isToRemoveRow = declinedValues.includes(rowValue);

        if (isToRemoveRow) removeRow();
      }
    }

    clearedValuesByColumns.unshift(headerRowNamesToReplaceHeaderRow);

    return clearedValuesByColumns;
  }
}
