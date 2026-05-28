const aliases = {

  // 일반암
  "암": "generalCancer",
  "일반암": "generalCancer",
  "암진단": "generalCancer",
  "암진단비": "generalCancer",

  // 고액암
  "고액암": "highCancer",

  // 유사암
  "유사암": "smallCancer",
  "소액암": "smallCancer",

  // 재진단
  "재진단암": "recurrenceCancer",
  "전이암": "recurrenceCancer",

  // 항암
  "항암": "antiCancer",
  "방사선": "antiCancer",
  "약물": "antiCancer",

  // 카티
  "카티": "carT",
  "표적항암": "carT",

  // 암수술
  "암수술": "cancerSurgery",

  // 뇌
  "뇌경색": "brainDiagnosis",
  "뇌출혈": "brainDiagnosis",
  "뇌혈관": "brainBloodVessel",

  // 심장
  "급성심근경색": "heartAttack",
  "심혈관": "cardioDiagnosis",
  "허혈성심장질환": "cardioDiagnosis",

  // 수술
  "질병수술": "diseaseSurgery",
  "상해수술": "accidentSurgery",

  // 입원
  "입원일당": "diseaseHospital",

  // 실손
  "실손": "actualDiseaseIn",

  // 골절
  "골절": "fractureDiagnosis",

  // 화상
  "화상": "burnDiagnosis",

  // 응급실
  "응급실": "emergency",

  // 운전자
  "변호사선임": "lawyerFee",
  "교통사고처리지원금": "trafficSupport"
};

const rowMap = {

  // 사망
  generalDeath: 8,
  accidentDeath: 9,
  diseaseDeath: 10,

  // 후유장해
  accidentDisability: 12,
  diseaseDisability: 13,

  // 암
  highCancer: 15,
  generalCancer: 16,
  smallCancer: 17,
  recurrenceCancer: 18,
  antiCancer: 19,
  cancerSurgery: 20,
  cancerPatient: 21,
  carT: 22,
  cancerTreatment: 23,

  // 뇌
  brainStroke: 25,
  brainDiagnosis: 26,
  brainBloodVessel: 27,
  brainSurgery: 28,

  // 심장
  heartAttack: 30,
  cardioDiagnosis: 31,
  cardioSurgery: 32,

  // 산정특례
  specialCaseOnce: 34,
  specialCaseYear: 35,

  // 상해수술
  accidentSurgery: 37,
  accidentSurgeryGrade: 38,

  // 질병수술
  diseaseSurgery: 40,
  diseaseSurgeryGrade: 41,
  majorSurgery: 42,

  // 입원일당
  accidentHospital: 45,
  diseaseHospital: 46,
  oneRoom: 47,
  multiRoom: 48,
  caregiver: 49,
  nursingCare: 50,

  // 실손
  actualAccidentIn: 52,
  actualAccidentOut: 53,
  actualDiseaseIn: 54,
  actualDiseaseOut: 55,
  nonBenefit: 56,

  // 골절
  fractureDiagnosis: 58,
  fractureSurgery: 59,
  castTreatment: 60,

  // 화상
  burnDiagnosis: 62,
  burnSurgery: 63,

  // 응급실
  emergency: 65,

  // 배상
  liability: 67,

  // 운전자
  trafficSupport: 69,
  lawyerFee: 70,
  licenseSupport: 71,
  loanSupport: 72,
  carAccident: 73
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

    for (const key in aliases) {

  const regex =
    new RegExp(`${key}\\s*([\\d,]+억?|[\\d,]+)`);

  const matched = line.match(regex);

  if (matched) {

    if (!matched[1].includes("억")) {

      result[aliases[key]] =
        Number(
          matched[1].replace(/,/g, "")
        ).toLocaleString();

        } else {

      result[aliases[key]] =
        matched[1];
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

  // 고객 정보 입력
  XLSX.utils.sheet_add_aoa(sheet, [[customerName]], { origin: 'Q1' });
  XLSX.utils.sheet_add_aoa(sheet, [[customerAge]], { origin: 'R1' });
  XLSX.utils.sheet_add_aoa(sheet, [[customerGender]], { origin: 'S1' });

  // 2. 보험 데이터 입력
  // 2칸~9칸 시트의 컬럼 배열을 정의하는 columnMap을 그대로 사용하되, 
  // 여기서는 단순히 보험 개수에 맞는 열(F, G, H...)을 순서대로 할당합니다.
 const insuranceCount = insuranceData.length;
 const columns = columnMap[`${insuranceCount}칸`];

  if (!columns) {
    alert("보험 개수는 2~9개까지만 가능합니다.");
    return;
  }

  insuranceData.forEach((insurance, index) => {
    const column = columns[index]; // 0번째 보험은 F열, 1번째는 G열...
    for (const key in insurance) {
      const row = rowMap[key];
      if (!row) continue;
      
      // Date 시트의 특정 좌표(예: F14, G14)에 값 입력
      XLSX.utils.sheet_add_aoa(sheet, [[insurance[key]]], { origin: column + row });
    }
  });

  // 3. 파일 저장
  XLSX.writeFile(workbook, `${customerName}_보장분석.xlsx`, { 
    bookType: 'xlsx', 
    compression: true 
  });
}
