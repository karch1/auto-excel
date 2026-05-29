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
  const fileCustomerName = document.getElementById("fileCustomerName").value.trim() || "고객";
  const excelCustomerName = document.getElementById("excelCustomerName").value.trim() || "고객";
  const customerAge = document.getElementById("customerAge").value.trim() || "0";
  const customerGender = document.getElementById("customerGender").value.trim() || "성별";

  const insuranceData = [];
  const insuranceMetaData = [];

  for (let i = 1; i <= 8; i++) {
    const text = document.getElementById(`insurance${i}`).value.trim();
    if (text) {
      insuranceData.push(parseCoverage(text));
      insuranceMetaData.push({
        payPeriod: document.getElementById(`payPeriod${i}`).value.trim(),
        company: document.getElementById(`company${i}`).value.trim(),
        product: document.getElementById(`product${i}`).value.trim(),
        joinDate: document.getElementById(`joinDate${i}`).value.trim(),
        monthlyFee: document.getElementById(`monthlyFee${i}`).value.trim()
      });
    }
  }

  const response = await fetch("./excel/template.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const sheet = workbook.getWorksheet("Date");

  if (!sheet) {
    alert("엑셀 파일에 'Date' 시트가 없습니다.");
    return;
  }

  sheet.getCell("Q1").value = excelCustomerName;
  sheet.getCell("R1").value = Number(customerAge);
  sheet.getCell("S1").value = customerGender;

  const insuranceCount = insuranceData.length;
  const columns = columnMap[`${insuranceCount}칸`];

  if (!columns) {
    alert("보험 개수는 2~9개까지만 가능합니다.");
    return;
  }

  // 통합 루프: 메타데이터 + 보장데이터를 한 번에 처리
  insuranceMetaData.forEach((meta, index) => {
    const column = columns[index];

   // 1. 메타 데이터 입력
sheet.getCell(`${column}3`).value = meta.payPeriod;
sheet.getCell(`${column}4`).value = meta.company;

// 상품명 셀
const productCell = sheet.getCell(`${column}5`);
const isRenewProduct = meta.product.includes("갱");
productCell.value = meta.product.replace(/\(갱\)/g, "").replace(/갱/g, "").trim();

// 스타일 객체를 가져와 수정 후 다시 적용 (기존 정렬 유지)
let productStyle = { ...productCell.style };
productStyle.alignment = { ...productStyle.alignment, shrinkToFit: true, vertical: 'middle', horizontal: 'center' };
if (isRenewProduct) {
    productStyle.font = { ...productStyle.font, color: { argb: "FF00B0F0" } };
}
productCell.style = productStyle;

sheet.getCell(`${column}6`).value = meta.joinDate;

// 납입보험료 셀
const feeCell = sheet.getCell(`${column}7`);
feeCell.value = Number(meta.monthlyFee.replace(/[^0-9]/g, ""));
feeCell.numFmt = '#,##0"원"';

// 스타일 객체를 가져와 수정 후 다시 적용
let feeStyle = { ...feeCell.style };
feeStyle.font = { ...feeStyle.font, color: { argb: "FFFF0000" }, name: 'Arial', size: 11 };
feeStyle.alignment = { ...feeStyle.alignment, vertical: 'middle', horizontal: 'center' };
feeCell.style = feeStyle;
    // 2. 보장 데이터 입력
    const currentInsurance = insuranceData[index];
    if (currentInsurance) {
      for (const key in currentInsurance) {
        const row = rowMap[key];
        if (!row) continue;
        const rawVal = String(currentInsurance[key]);
        const isRenew = rawVal.includes("갱");
        const val = rawVal.replace(/\(갱\)/g, "").replace(/갱/g, "").trim();
        const cell = sheet.getCell(column + row);
        const cleanNumber = Number(val.replace(/,/g, ""));

        if (val.includes("억") || isNaN(cleanNumber)) {
          cell.value = val;
        } else {
          cell.value = cleanNumber;
          cell.numFmt = '#,##0';
        }
        if (isRenew) cell.font = { color: { argb: "FF00B0F0" } };
      }
    }
  });

  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${fileCustomerName}_보장분석.xlsx`;
  link.click();
}

// textarea 자동 높이
document.querySelectorAll("textarea").forEach(textarea => {

  // 최초 실행
  textarea.style.height = "auto";
  textarea.style.height = textarea.scrollHeight + "px";

  // 입력 시 자동 증가
  textarea.addEventListener("input", () => {

    textarea.style.height = "auto";
    textarea.style.height = textarea.scrollHeight + "px";

  });

});

for (let i = 1; i <= 8; i++) {

  // 가입연도 자동 포맷
  const joinDateInput =
  document.getElementById(`joinDate${i}`);

  joinDateInput.addEventListener("input", e => {

    let value =
    e.target.value.replace(/[^0-9]/g, "");

    if (value.length > 8) {
      value = value.slice(0, 8);
    }

    if (value.length >= 5) {

      value =
        value.slice(0,4) + "." +
        value.slice(4,6) + "." +
        value.slice(6,8);

    }

    e.target.value = value;

  });

  // 납입보험료 자동 포맷
  const monthlyFeeInput =
  document.getElementById(`monthlyFee${i}`);

  monthlyFeeInput.addEventListener("input", e => {

    let value =
    e.target.value.replace(/[^0-9]/g, "");

    if (!value) {
      e.target.value = "";
      return;
    }

    e.target.value =
      Number(value).toLocaleString() + "원";

  });

}
