const getColumnLetter = (colNumber) => {
  let letter = "";
  while (colNumber > 0) {
    const remainder = (colNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    colNumber = Math.floor((colNumber - 1) / 26);
  }
  return letter;
};

const setBorder = (cell, style = "thin") => {
  cell.border = {
    top: { style },
    left: { style },
    bottom: { style },
    right: { style },
  };
};

const mergeAndStyle = (worksheet, range, value, options = {}) => {
  worksheet.mergeCells(range);
  const cell = worksheet.getCell(range.split(":")[0]);
  cell.value = value;
  if (options.font) cell.font = options.font;
  if (options.alignment) cell.alignment = options.alignment;
  if (options.fill) cell.fill = options.fill;
  if (options.border) setBorder(cell, options.border);
  return cell;
};

const getNamaPendekKapal = (namaKapal) => {
  let n = String(namaKapal).trim();
  if (n.toUpperCase().startsWith("KMP.")) n = n.substring(4).trim();
  return n
    .split(" ")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
};

module.exports = {
  getColumnLetter,
  setBorder,
  mergeAndStyle,
  getNamaPendekKapal,
};
