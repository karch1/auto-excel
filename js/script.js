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

async function generateExcel() {
  const customerName = document.getElementById("customerName").value.trim() || "고객";
  const customerAge = document.getElementById("customerAge").value.trim() || "0";
  const customerGender = document.getElementById("customerGender").value.trim() || "성별";

  const insuranceData = [];
  for (let i = 1; i <= 8; i++) {
    const text = document.getElementById(`insurance${i}`).value.trim();
    if (text) insuranceData.push(parseCoverage(text));
  }

  const response = await fetch("./excel/template.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  // 핵심: 이제 시트 이름(2칸, 3칸...)을 찾지 않고 무조건 'Date' 시트에 씁니다.
  const sheet = workbook.Sheets['Date']; 
  
  if (!sheet) {
    alert("엑셀 파일에 'Date' 시트가 없습니다.");
    return;
  }

  // 1. 고객 정보 입력 (Date 시트의 A1, A2, C3 셀에 씀)
  XLSX.utils.sheet_add_aoa(sheet, [[customerName]], { origin: 'A1' });
  XLSX.utils.sheet_add_aoa(sheet, [[customerAge]], { origin: 'A2' });
  XLSX.utils.sheet_add_aoa(sheet, [[customerGender]], { origin: 'C3' });

  // 2. 보험 데이터 입력
  // 2칸~9칸 시트의 컬럼 배열을 정의하는 columnMap을 그대로 사용하되, 
  // 여기서는 단순히 보험 개수에 맞는 열(F, G, H...)을 순서대로 할당합니다.
  const insuranceCount = insuranceData.length;
  const columns = ["F", "G", "H", "I", "J", "K", "L", "M", "N"];

  insuranceData.forEach((insurance, index) => {
    const column = columns[index]; // 0번째 보험은 F열, 1번째는 G열...
    for (const key in insurance) {
      const row = rowMap[key];
      if (!row) continue;
      
      // Date 시트의 특정 좌표(예: F14, G14)에 값 입력
      XLSX.utils.sheet_add_aoa(sheet, [[insurance[key].toLocaleString()]], { origin: column + row });
    }
  });

  // 3. 파일 저장
  XLSX.writeFile(workbook, `${customerName}_보장분석.xlsx`, { 
    bookType: 'xlsx', 
    compression: true 
  });
}
