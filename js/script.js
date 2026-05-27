const aliases = {

  "암": "generalCancer",
  "소액암": "smallCancer",
  "유사암": "smallCancer",

  "카티": "carT",

  "심혈관수술": "cardio",
  "심혈관진단": "cardio",

  "뇌혈관진단": "brain"
};

const rowMap = {

  generalCancer: 14,
  smallCancer: 15,
  carT: 20,
  cardio: 27,
  brain: 24
};

const columnMap = {

  "2칸": ["F","G"],
  "3칸": ["F","G","H"],
  "4칸": ["F","G","H","I"],
  "5칸": ["F","G","H","I","J"],
  "6칸": ["F","G","H","I","J","K"],
  "7칸": ["F","G","H","I","J","K","L"],
  "8칸": ["F","G","H","I","J","K","L","M"],
  "9칸": ["F","G","H","I","J","K","L","M","N"]

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

        const matched =
          line.match(/(\d[\d,]*)/);

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

  const customerName =
    document
      .getElementById("customerName")
      .value
      .trim() || "고객";

  const insuranceData = [];

  for(let i=1; i<=8; i++){

    const text =
      document
        .getElementById(`insurance${i}`)
        .value
        .trim();

    if(text){

      insuranceData.push(
        parseCoverage(text)
      );
    }
  }

  const insuranceCount =
    insuranceData.length;

  const sheetName =
    `${insuranceCount}칸`;

  const response =
    await fetch("./excel/template.xlsx");

  const arrayBuffer =
    await response.arrayBuffer();

  const workbook =
    XLSX.read(arrayBuffer, {
      type: "array"
    });

  console.log(workbook.SheetNames);

  const sheet =
    workbook.Sheets[sheetName];

  if(!sheet){

    alert(`${sheetName} 시트를 찾을 수 없습니다.`);
    return;
  }

  const columns =
    columnMap[sheetName];

  insuranceData.forEach((insurance, index) => {

    const column =
      columns[index];

    for(const key in insurance){

      const row =
        rowMap[key];

      if(!row) continue;

      const cellAddress =
        column + row;

      if(sheet[cellAddress]){

        sheet[cellAddress].v =
          insurance[key]
            .toLocaleString();

      }else{

        XLSX.utils.sheet_add_aoa(
          sheet,
          [[insurance[key].toLocaleString()]],
          { origin: cellAddress }
        );
      }
    }
  });

    XLSX.writeFile(workbook, `${customerName}_보장분석.xlsx`, { 
      bookType: 'xlsx', 
      type: 'binary',
      compression: true // 압축 옵션 추가
  });
}
