import baseConfig from "ts-builds/eslint-functype"

export default [
  ...baseConfig,
  {
    // Calibration — @sapientsai/ms-graph-extract wraps three imperative parser libraries (mammoth,
    // unpdf, exceljs). Each returns by throwing, so try/catch-to-Either is the irreducible shape
    // here, and exceljs's eachSheet/eachRow visitors accumulate into arrays because that is the only
    // traversal API it exposes. prefer-fold fires on `v == null ? "" : String(v)`, which is a raw
    // nullable from exceljs's cell values, not a functype Option. The functional surface is the
    // Either this package returns; prefer-either stays ON to keep it that way.
    //
    // Carried over from packages/graph's `src/extract/**` block, where this code previously lived.
    files: ["src/**/*.ts"],
    rules: {
      "functype/prefer-try": "off",
      "functype/prefer-fold": "off",
      "functype/no-imperative-loops": "off",
    },
  },
]
