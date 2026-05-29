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
    // 1. 사망 (8~10행)
    generalDeath: 8, accidentDeath: 9, diseaseDeath: 10,
    // 2. 후유장해 (11~12행)
    accidentDisability: 11, diseaseDisability: 12,
    // 3. 3대 진단·수술 (암 13~21, 뇌 22~25, 심장 26~28)
    highCancer: 13, generalCancer: 14, smallCancer: 15, recurrenceCancer: 16,
    antiCancer: 17, cancerSurgery: 18, cancerPatient: 19, carT: 20, cancerTreatment: 21,
    brainDiagnosis: 22, brainStroke: 23, brainBloodVessel: 24, brainSurgery: 25,
    heartAttack: 26, cardioDiagnosis: 27, cardioSurgery: 28,
    // 4. 산정특례 (29~30행)
    specialCaseOnce: 29, specialCaseYear: 30,
    // 5. 수술 (상해 31~32, 질병 33~36)
    accidentSurgery: 31, accidentSurgeryGrade: 32,
    diseaseSurgery: 33, diseaseSurgeryGrade: 34, majorSurgery: 35,
    // 6. 입원일당 (37~42행)
    accidentHospital: 37, diseaseHospital: 38, oneRoom: 39, 
    multiRoom: 40, caregiver: 41, nursingCare: 42,
    // 7. 실손의료비 (43~47행)
    actualAccidentIn: 43, actualAccidentOut: 44, 
    actualDiseaseIn: 45, actualDiseaseOut: 46, nonBenefit: 47,
    // 8. 골절 (48~50행)
    fractureDiagnosis: 48, fractureSurgery: 49, castTreatment: 50,
    // 9. 화상 (51~52행)
    burnDiagnosis: 51, burnSurgery: 52,
    // 10. 응급실 (53행)
    emergency: 53,
    // 11. 배상 (54행)
    liability: 54,
    // 12. 운전자보험 (55~60행)
    trafficSupport: 55, lawyerFee: 56, licenseSupport: 57, 
    loanSupport: 58, carAccident: 59
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
  new RegExp(`${key}[^\\d]*([\\d,.]+(?:억|천만|백만|만)?(?:원)?)`);

  const matched = line.match(regex);

    if (matched) {

      const value = matched[1]
        .replace(/원/g, "")
        .trim();

      if (
        value.includes("억") ||
        value.includes("만")
        ) {
  
        result[aliases[key]] = value;
    
    } else {

            result[aliases[key]] =
            Number(
            value.replace(/,/g, "")
          ).toLocaleString();
        }
      }
    }
  });

  return result;
}

async function generateExcel() {
  // 다운로드이름
  const fileCustomerName =
  document.getElementById("fileCustomerName").value.trim() || "고객";
  // 안에 고객 이름
  const excelCustomerName =
  document.getElementById("excelCustomerName").value.trim() || "고객";
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
  if (!sheet["Q1"]) sheet["Q1"] = {};
  if (!sheet["R1"]) sheet["R1"] = {};
  if (!sheet["S1"]) sheet["S1"] = {};

sheet["Q1"].v = excelCustomerName;
sheet["R1"].v = customerAge;
sheet["S1"].v = customerGender;

sheet["Q1"].t = "s";
sheet["R1"].t = "n";
sheet["S1"].t = "s";

  // 2. 보험 데이터 입력
  // 2칸~9칸 시트의 컬럼 배열을 정의하는 columnMap을 그대로 사용하되, 
  // 여기서는 단순히 보험 개수에 맞는 열(F, G, H...)을 순서대로 할당합니다.
 const insuranceCount = insuranceData.length;
 const columns = columnMap[`${insuranceCount}칸`];

  if (!columns) {
    alert("보험 개수는 2~9개까지만 가능합니다.");
    return;
  }

  // 2. 보험 데이터 입력 부분 수정
  insuranceData.forEach((insurance, index) => {
    const column = columns[index]; // F, G, H...
    
    for (const key in insurance) {
      const row = rowMap[key];
      if (!row) continue;
      
     const address = column + row;
const val = String(insurance[key]);

      // 셀없으면 생성
if (!sheet[address]) {
  sheet[address] = {};
}

      // 
const cleanNumber = Number(
  val.replace(/,/g, "")
);

if (
  val.includes("억") ||
  isNaN(cleanNumber)
) {
  sheet[address].v = val;
  sheet[address].t = "s";
} else {
  sheet[address].v = cleanNumber;
  sheet[address].t = "n";
}
      }
  });

  // 3. 파일 저장
  XLSX.writeFile(workbook,`${fileCustomerName}_보장분석.xlsx`, { 
    bookType: 'xlsx', 
    compression: true 
  });
}
