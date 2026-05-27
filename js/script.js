const aliases = {
  "암": "generalCancer",
  "소액암": "smallCancer",
  "카티": "carT"
};

const rowMap = {
  generalCancer: 20,
  smallCancer: 22,
  carT: 24
};

document
  .getElementById("downloadBtn")
  .addEventListener("click", generateExcel);

function parseCoverage(text){

  const result = {};

  const lines = text.split("\n");

  lines.forEach(line => {

    line = line.trim();

    for(const key in aliases){

      if(line.includes(key)){

        const matched = line.match(/(\d[\d,]*)/);

        if(matched){

          const amount = Number(
            matched[1].replace(/,/g, "")
          );

          result[aliases[key]] = amount;
        }
      }
    }
  });

  return result;
}

async function generateExcel(){

  const insurance1Text =
    document.getElementById("insurance1").value;

  const parsedData =
    parseCoverage(insurance1Text);

  console.log(parsedData);

  const response =
    await fetch("./excel/template.xlsx");

  const arrayBuffer =
    await response.arrayBuffer();

  const workbook =
    XLSX.read(arrayBuffer, {
      type: "array"
    });

  const sheet =
    workbook.Sheets[workbook.SheetNames[0]];

  const column = "D";

  for(const key in parsedData){

    const row = rowMap[key];

    if(!row) continue;

    const cellAddress = column + row;

    sheet[cellAddress] = {
      t: "s",
      v: parsedData[key].toLocaleString()
    };
  }

  XLSX.writeFile(
    workbook,
    "보장분석.xlsx"
  );
}