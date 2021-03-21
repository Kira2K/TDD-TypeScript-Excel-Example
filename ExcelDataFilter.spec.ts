import { ExcelDataFilter } from "./ExcelDataFilter";

describe("Testing ExcelDataFilter Module ", () => {
  const module = new ExcelDataFilter();
  const headerRowNamesToReplaceHeaderRow = [
    "TestHeader1",
    true,
    1,
    null,
    undefined,
    "",
  ];
  const declinedHeadersColumnValuesArr: ExcelDataFilter.getFilteredRowsByHeaderColumnValuesArg["declinedHeadersColumnValuesArr"] = [
    {
      columnRawValues: [[]],
      declinedColumnValues: [12],
      index: 0,
      name: "testHeader1",
    },
    {
      columnRawValues: [[]],
      declinedColumnValues: [22, 23],
      index: 1,
      name: "testHeader1",
    },
  ];
  const rawValues = [
    ["test __ header 1", "true", "", 4, null, undefined],
    [],
    [11, 21, null],
    [12, 22, 31],
    [undefined, undefined, undefined],
    [13, 23, undefined, null],
  ];
  const filteredValuesByHeaders = [
    headerRowNamesToReplaceHeaderRow,
    rawValues[2],
  ];
  const skipRowsQuantity = 1;
  it("getFilteredValuesInColumnsByHeaders test", async () => {
    const values = module.getFilteredRowsByHeaderColumnValues({
      declinedHeadersColumnValuesArr,
      headerRowNamesToReplaceHeaderRow,
      values: rawValues,
      skipRowsQuantity,
    });
    expect(values).toEqual(filteredValuesByHeaders);
  });
});
