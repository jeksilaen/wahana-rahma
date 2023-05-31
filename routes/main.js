require("dotenv").config();

const express = require("express");
const reader = require("xlsx");
const ExcelJS = require("exceljs");
const router = express.Router();
const nodemailer = require("nodemailer");
const fs = require("fs");

router.get("/", (req, res) => {
  res.sendFile("index.html", { root: "./" });
});

router.get("/sucess", (req, res) => {
  res.sendFile("sucess.html", { root: "./" });
});

router.post("/", async (req, res) => {
  const {
    engName,
    dob,
    height,
    overallSize,
    nextOfKin,
    relation,
    engAddress,
    idNo,
    chiName,
    age,
    weight,
    safeShoeSize,
    chiAddress,
    postCode,
    appliedRank,
    pob,
    bmi,
    maritalStatus,
    noOfChild,
    mobile,
    emergencyContact,
    hometownAirport,
    availableTime,
    gradePassport,
    noPassport,
    tglAwalPassport,
    tempatPassport,
    tglAkhirPassport,
    gradeSeafarer,
    noSeafarer,
    tglAwalSeafarer,
    tempatSeafarer,
    tglAkhirSeafarer,
    gradeSeaman,
    noSeaman,
    tglAwalSeaman,
    tempatSeaman,
    tglAkhirSeaman,
    gradeCOC,
    noCOC,
    tglAwalCOC,
    tempatCOC,
    tglAkhirCOC,
    gradeGMDSS,
    noGMDSS,
    tglAwalGMDSS,
    tempatGMDSS,
    tglAkhirGMDSS,
    gradeUSVISA,
    noUSVISA,
    tglAwalUSVISA,
    tempatUSVISA,
    tglAkhirUSVISA,
    gradeMedical,
    noMedical,
    tglAwalMedical,
    tempatMedical,
    tglAkhirMedical,
    gradeYellow,
    noYellow,
    tglAwalYellow,
    tempatYellow,
    tglAkhirYellow,
    gradeCholera,
    noCholera,
    tglAwalCholera,
    tempatCholera,
    tglAkhirCholera,
    noFamiliarization,
    tglAwalFamiliarization,
    tempatFamiliarization,
    tglAkhirFamiliarization,
    noProficiencySurvival,
    tglAwalProficiencySurvival,
    tempatProficiencySurvival,
    tglAkhirProficiencySurvival,
    noAdvancedFirefighting,
    tglAwalAdvancedFirefighting,
    tempatAdvancedFirefighting,
    tglAkhirAdvancedFirefighting,
    noProficiencyMedical,
    tglAwalProficiencyMedical,
    tempatProficiencyMedical,
    tglAkhirProficiencyMedical,
    noMedicalCare,
    tglAwalMedicalCare,
    tempatMedicalCare,
    tglAkhirMedicalCare,
    noShipSecurity,
    tglAwalShipSecurity,
    tempatShipSecurity,
    tglAkhirShipSecurity,
    noBridgeTeam,
    tglAwalBridgeTeam,
    tempatBridgeTeam,
    tglAkhirBridgeTeam,
    noShipHandling,
    tglAwalShipHandling,
    tempatShipHandling,
    tglAkhirShipHandling,
    noSecurityAwareness,
    tglAwalSecurityAwareness,
    tempatSecurityAwareness,
    tglAkhirSecurityAwareness,
    noSeafarersDesignated,
    tglAwalSeafarersDesignated,
    tempatSeafarersDesignated,
    tglAkhirSeafarersDesignated,
    noECDIS,
    tglAwalECDIS,
    tempatECDIS,
    tglAkhirECDIS,
    noECDISType,
    tglAwalECDISType,
    tempatECDISType,
    tglAkhirECDISType,
    noAdvancedTrainingOil,
    tglAwalAdvancedTrainingOil,
    tempatAdvancedTrainingOil,
    tglAkhirAdvancedTrainingOil,
    noAdvancedTrainingChemical,
    tglAwalAdvancedTrainingChemical,
    tempatAdvancedTrainingChemical,
    tglAkhirAdvancedTrainingChemical,
    noBasicTrainingOil,
    tglAwalBasicTrainingOil,
    tempatBasicTrainingOil,
    tglAkhirBasicTrainingOil,
    gradePanama,
    noPanamaLicense,
    tglAwalPanamaLicense,
    tglAkhirPanamaLicense,
    noPanamaGMDSS,
    tglAwalPanamaGMDSS,
    tglAkhirPanamaGMDSS,
    noPanamaOther,
    tglPanamaOther,
    tglAkhirPanamaOther,
    gradeHongkong,
    noHongkongLicense,
    tglAwalHongkongLicense,
    tglAkhirHongkongLicense,
    noHongkongGMDSS,
    tglAwalHongkongGMDSS,
    tglAkhirHongkongGMDSS,
    noHongkongOther,
    tglHongkongOther,
    tglAkhirHongkongOther,
    gradeTuvalu,
    noTuvaluLicense,
    tglAwalTuvaluLicense,
    tglAkhirTuvaluLicense,
    noTuvaluGMDSS,
    tglAwalTuvaluGMDSS,
    tglAkhirTuvaluGMDSS,
    noTuvaluOther,
    tglTuvaluOther,
    tglAkhirTuvaluOther,
    nameOfInstitution,
    cityPlusCountry,
    qualificationFrom,
    qualificationTo,
    qualificationObtained,
    seaServiceCompany1,
    seaServiceRank1,
    seaServiceVessel1,
    seaServiceType1,
    seaServiceFlag1,
    seaServiceGRT1,
    seaServiceEngine1,
    seaServiceFrom1,
    seaServiceTo1,
    seaServiceMOs1,
    seaServiceReason1,
    seaServiceCompany2,
    seaServiceRank2,
    seaServiceVessel2,
    seaServiceType2,
    seaServiceFlag2,
    seaServiceGRT2,
    seaServiceEngine2,
    seaServiceFrom2,
    seaServiceTo2,
    seaServiceMOs2,
    seaServiceReason2,
    seaServiceCompany3,
    seaServiceRank3,
    seaServiceVessel3,
    seaServiceType3,
    seaServiceFlag3,
    seaServiceGRT3,
    seaServiceEngine3,
    seaServiceFrom3,
    seaServiceTo3,
    seaServiceMOs3,
    seaServiceReason3,
    seaServiceCompany4,
    seaServiceRank4,
    seaServiceVessel4,
    seaServiceType4,
    seaServiceFlag4,
    seaServiceGRT4,
    seaServiceEngine4,
    seaServiceFrom4,
    seaServiceTo4,
    seaServiceMOs4,
    seaServiceReason4,
    seaServiceCompany5,
    seaServiceRank5,
    seaServiceVessel5,
    seaServiceType5,
    seaServiceFlag5,
    seaServiceGRT5,
    seaServiceEngine5,
    seaServiceFrom5,
    seaServiceTo5,
    seaServiceMOs5,
    seaServiceReason5,
  } = req.body;

  const filePath = __dirname + "/template.xlsx";

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet(1);

  // Personal Particulars
  worksheet.getCell("B6").value = engName;
  worksheet.getCell("B7").value = dob;
  worksheet.getCell("B8").value = height;
  worksheet.getCell("B9").value = overallSize;
  worksheet.getCell("B10").value = nextOfKin;
  worksheet.getCell("B10").value = nextOfKin;
  worksheet.getCell("F10").value = relation;
  worksheet.getCell("A12").value = engAddress;
  worksheet.getCell("B14").value = idNo;

  worksheet.getCell("F6").value = chiName;
  worksheet.getCell("F7").value = age;
  worksheet.getCell("F8").value = weight;
  worksheet.getCell("F9").value = safeShoeSize;
  worksheet.getCell("D12").value = chiAddress;
  worksheet.getCell("F14").value = postCode;

  worksheet.getCell("J6").value = appliedRank;
  worksheet.getCell("J7").value = pob;
  worksheet.getCell("J8").value = bmi;
  worksheet.getCell("J9").value = maritalStatus;
  worksheet.getCell("J10").value = noOfChild;
  worksheet.getCell("J11").value = mobile;
  worksheet.getCell("J12").value = emergencyContact;
  worksheet.getCell("J13").value = hometownAirport;
  worksheet.getCell("J14").value = availableTime;

  // Certificates
  worksheet.getCell("B18").value = gradePassport;
  worksheet.getCell("B19").value = gradeSeafarer;
  worksheet.getCell("B20").value = gradeSeaman;
  worksheet.getCell("B21").value = gradeCOC;
  worksheet.getCell("B22").value = gradeGMDSS;
  worksheet.getCell("B23").value = gradeUSVISA;
  worksheet.getCell("B24").value = gradeMedical;
  worksheet.getCell("B25").value = gradeYellow;
  worksheet.getCell("B26").value = gradeCholera;

  worksheet.getCell("C18").value = noPassport;
  worksheet.getCell("C19").value = noSeafarer;
  worksheet.getCell("C20").value = noSeaman;
  worksheet.getCell("C21").value = noCOC;
  worksheet.getCell("C22").value = noGMDSS;
  worksheet.getCell("C23").value = noUSVISA;
  worksheet.getCell("C24").value = noMedical;
  worksheet.getCell("C25").value = noYellow;
  worksheet.getCell("C26").value = noCholera;

  worksheet.getCell("F18").value = tglAwalPassport;
  worksheet.getCell("F19").value = tglAwalSeafarer;
  worksheet.getCell("F20").value = tglAwalSeaman;
  worksheet.getCell("F21").value = tglAwalCOC;
  worksheet.getCell("F22").value = tglAwalGMDSS;
  worksheet.getCell("F23").value = tglAwalUSVISA;
  worksheet.getCell("F24").value = tglAwalMedical;
  worksheet.getCell("F25").value = tglAwalYellow;
  worksheet.getCell("F26").value = tglAwalCholera;

  worksheet.getCell("H18").value = tempatPassport;
  worksheet.getCell("H19").value = tempatSeafarer;
  worksheet.getCell("H20").value = tempatSeaman;
  worksheet.getCell("H21").value = tempatCOC;
  worksheet.getCell("H22").value = tempatGMDSS;
  worksheet.getCell("H23").value = tempatUSVISA;
  worksheet.getCell("H24").value = tempatMedical;
  worksheet.getCell("H25").value = tempatYellow;
  worksheet.getCell("H26").value = tempatCholera;

  worksheet.getCell("K18").value = tglAkhirPassport;
  worksheet.getCell("K19").value = tglAkhirSeafarer;
  worksheet.getCell("K20").value = tglAkhirSeaman;
  worksheet.getCell("K21").value = tglAkhirCOC;
  worksheet.getCell("K22").value = tglAkhirGMDSS;
  worksheet.getCell("K23").value = tglAkhirUSVISA;
  worksheet.getCell("K24").value = tglAkhirMedical;
  worksheet.getCell("K25").value = tglAkhirYellow;
  worksheet.getCell("K26").value = tglAkhirCholera;

  // STCW Certificates
  worksheet.getCell("B29").value = noFamiliarization;
  worksheet.getCell("B30").value = noProficiencySurvival;
  worksheet.getCell("B31").value = noAdvancedFirefighting;
  worksheet.getCell("B32").value = noProficiencyMedical;
  worksheet.getCell("B33").value = noMedicalCare;
  worksheet.getCell("B34").value = noShipSecurity;
  worksheet.getCell("B35").value = noBridgeTeam;
  worksheet.getCell("B36").value = noShipHandling;
  worksheet.getCell("B37").value = noSecurityAwareness;
  worksheet.getCell("B38").value = noSeafarersDesignated;
  worksheet.getCell("B39").value = noECDIS;
  worksheet.getCell("B40").value = noECDISType;
  worksheet.getCell("B41").value = noAdvancedTrainingOil;
  worksheet.getCell("B42").value = noAdvancedTrainingChemical;
  worksheet.getCell("B43").value = noBasicTrainingOil;

  worksheet.getCell("C29").value = tglAwalFamiliarization;
  worksheet.getCell("C30").value = tglAwalProficiencySurvival;
  worksheet.getCell("C31").value = tglAwalAdvancedFirefighting;
  worksheet.getCell("C32").value = tglAwalProficiencyMedical;
  worksheet.getCell("C33").value = tglAwalMedicalCare;
  worksheet.getCell("C34").value = tglAwalShipSecurity;
  worksheet.getCell("C35").value = tglAwalBridgeTeam;
  worksheet.getCell("C36").value = tglAwalShipHandling;
  worksheet.getCell("C37").value = tglAwalSecurityAwareness;
  worksheet.getCell("C38").value = tglAwalSeafarersDesignated;
  worksheet.getCell("C39").value = tglAwalECDIS;
  worksheet.getCell("C40").value = tglAwalECDISType;
  worksheet.getCell("C41").value = tglAwalAdvancedTrainingOil;
  worksheet.getCell("C42").value = tglAwalAdvancedTrainingChemical;
  worksheet.getCell("C43").value = tglAwalBasicTrainingOil;

  worksheet.getCell("D29").value = tempatFamiliarization;
  worksheet.getCell("D30").value = tempatProficiencySurvival;
  worksheet.getCell("D31").value = tempatAdvancedFirefighting;
  worksheet.getCell("D32").value = tempatProficiencyMedical;
  worksheet.getCell("D33").value = tempatMedicalCare;
  worksheet.getCell("D34").value = tempatShipSecurity;
  worksheet.getCell("D35").value = tempatBridgeTeam;
  worksheet.getCell("D36").value = tempatShipHandling;
  worksheet.getCell("D37").value = tempatSecurityAwareness;
  worksheet.getCell("D38").value = tempatSeafarersDesignated;
  worksheet.getCell("D39").value = tempatECDIS;
  worksheet.getCell("D40").value = tempatECDISType;
  worksheet.getCell("D41").value = tempatAdvancedTrainingOil;
  worksheet.getCell("D42").value = tempatAdvancedTrainingChemical;
  worksheet.getCell("D43").value = tempatBasicTrainingOil;

  worksheet.getCell("E29").value = tglAkhirFamiliarization;
  worksheet.getCell("E30").value = tglAkhirProficiencySurvival;
  worksheet.getCell("E31").value = tglAkhirAdvancedFirefighting;
  worksheet.getCell("E32").value = tglAkhirProficiencyMedical;
  worksheet.getCell("E33").value = tglAkhirMedicalCare;
  worksheet.getCell("E34").value = tglAkhirShipSecurity;
  worksheet.getCell("E35").value = tglAkhirBridgeTeam;
  worksheet.getCell("E36").value = tglAkhirShipHandling;
  worksheet.getCell("E37").value = tglAkhirSecurityAwareness;
  worksheet.getCell("E38").value = tglAkhirSeafarersDesignated;
  worksheet.getCell("E39").value = tglAkhirECDIS;
  worksheet.getCell("E40").value = tglAkhirECDISType;
  worksheet.getCell("E41").value = tglAkhirAdvancedTrainingOil;
  worksheet.getCell("E42").value = tglAkhirAdvancedTrainingChemical;
  worksheet.getCell("E43").value = tglAkhirBasicTrainingOil;

  // Equivalent Endorsements
  worksheet.getCell("B47").value = gradePanama;
  worksheet.getCell("B48").value = gradeHongkong;
  worksheet.getCell("B49").value = gradeTuvalu;

  worksheet.getCell("C47").value = noPanamaLicense;
  worksheet.getCell("C48").value = noHongkongLicense;
  worksheet.getCell("C49").value = noTuvaluLicense;

  worksheet.getCell("D47").value = tglAwalPanamaLicense;
  worksheet.getCell("D48").value = tglAwalHongkongLicense;
  worksheet.getCell("D49").value = tglAwalTuvaluLicense;

  worksheet.getCell("F47").value = tglAkhirPanamaLicense;
  worksheet.getCell("F48").value = tglAkhirHongkongLicense;
  worksheet.getCell("F49").value = tglAkhirTuvaluLicense;

  worksheet.getCell("G47").value = noPanamaGMDSS;
  worksheet.getCell("G48").value = noHongkongGMDSS;
  worksheet.getCell("G49").value = noTuvaluGMDSS;

  worksheet.getCell("H47").value = tglAwalPanamaGMDSS;
  worksheet.getCell("H48").value = tglAwalHongkongGMDSS;
  worksheet.getCell("H49").value = tglAwalTuvaluGMDSS;

  worksheet.getCell("J47").value = tglAkhirPanamaGMDSS;
  worksheet.getCell("J48").value = tglAkhirHongkongGMDSS;
  worksheet.getCell("J49").value = tglAkhirTuvaluGMDSS;

  worksheet.getCell("K47").value = noPanamaOther;
  worksheet.getCell("K48").value = noHongkongOther;
  worksheet.getCell("K49").value = noTuvaluOther;

  worksheet.getCell("L47").value = tglPanamaOther;
  worksheet.getCell("L48").value = tglHongkongOther;
  worksheet.getCell("L49").value = tglTuvaluOther;

  worksheet.getCell("F47").value = tglAkhirPanamaOther;
  worksheet.getCell("F48").value = tglAkhirHongkongOther;
  worksheet.getCell("F49").value = tglAkhirTuvaluOther;

  // Qualifications
  worksheet.getCell("A52").value = nameOfInstitution;
  worksheet.getCell("D52").value = cityPlusCountry;
  worksheet.getCell("G52").value = qualificationFrom;
  worksheet.getCell("I52").value = qualificationTo;
  worksheet.getCell("K52").value = qualificationObtained;

  // Sea Service Record In Last 5 Years
  worksheet.getCell("A55").value = seaServiceCompany1;
  worksheet.getCell("C55").value = seaServiceRank1;
  worksheet.getCell("D55").value = seaServiceVessel1;
  worksheet.getCell("E55").value = seaServiceType1;
  worksheet.getCell("F55").value = seaServiceFlag1;
  worksheet.getCell("G55").value = seaServiceGRT1;
  worksheet.getCell("H55").value = seaServiceEngine1;
  worksheet.getCell("I55").value = seaServiceFrom1;
  worksheet.getCell("J55").value = seaServiceTo1;
  worksheet.getCell("K55").value = seaServiceMOs1;
  worksheet.getCell("L55").value = seaServiceReason1;

  worksheet.getCell("A56").value = seaServiceCompany2;
  worksheet.getCell("C56").value = seaServiceRank2;
  worksheet.getCell("D56").value = seaServiceVessel2;
  worksheet.getCell("E56").value = seaServiceType2;
  worksheet.getCell("F56").value = seaServiceFlag2;
  worksheet.getCell("G56").value = seaServiceGRT2;
  worksheet.getCell("H56").value = seaServiceEngine2;
  worksheet.getCell("I56").value = seaServiceFrom2;
  worksheet.getCell("J56").value = seaServiceTo2;
  worksheet.getCell("K56").value = seaServiceMOs2;
  worksheet.getCell("L56").value = seaServiceReason2;

  worksheet.getCell("A57").value = seaServiceCompany3;
  worksheet.getCell("C57").value = seaServiceRank3;
  worksheet.getCell("D57").value = seaServiceVessel3;
  worksheet.getCell("E57").value = seaServiceType3;
  worksheet.getCell("F57").value = seaServiceFlag3;
  worksheet.getCell("G57").value = seaServiceGRT3;
  worksheet.getCell("H57").value = seaServiceEngine3;
  worksheet.getCell("I57").value = seaServiceFrom3;
  worksheet.getCell("J57").value = seaServiceTo3;
  worksheet.getCell("K57").value = seaServiceMOs3;
  worksheet.getCell("L57").value = seaServiceReason3;

  worksheet.getCell("A58").value = seaServiceCompany4;
  worksheet.getCell("C58").value = seaServiceRank4;
  worksheet.getCell("D58").value = seaServiceVessel4;
  worksheet.getCell("E58").value = seaServiceType4;
  worksheet.getCell("F58").value = seaServiceFlag4;
  worksheet.getCell("G58").value = seaServiceGRT4;
  worksheet.getCell("H58").value = seaServiceEngine4;
  worksheet.getCell("I58").value = seaServiceFrom4;
  worksheet.getCell("J58").value = seaServiceTo4;
  worksheet.getCell("K58").value = seaServiceMOs4;
  worksheet.getCell("L58").value = seaServiceReason4;

  worksheet.getCell("A59").value = seaServiceCompany5;
  worksheet.getCell("C59").value = seaServiceRank5;
  worksheet.getCell("D59").value = seaServiceVessel5;
  worksheet.getCell("E59").value = seaServiceType5;
  worksheet.getCell("F59").value = seaServiceFlag5;
  worksheet.getCell("G59").value = seaServiceGRT5;
  worksheet.getCell("H59").value = seaServiceEngine5;
  worksheet.getCell("I59").value = seaServiceFrom5;
  worksheet.getCell("J59").value = seaServiceTo5;
  worksheet.getCell("K59").value = seaServiceMOs5;
  worksheet.getCell("L59").value = seaServiceReason5;

  const outputPath = filePath.replace(".xlsx", "_modified.xlsx");
  await workbook.xlsx.writeFile(outputPath);
  await sendEmailWithAttachment();

  console.log("Changes saved to:", outputPath);

  res.sendFile("sucess.html", { root: "./" });
});

async function sendEmailWithAttachment() {
  // Read the XLSX file as a buffer
  const file = fs.readFileSync(__dirname + "/template_modified.xlsx");

  // Create a Nodemailer transporter
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: "ippo7707@gmail.com",
      pass: process.env.PASS,
    },
  });

  // Define email options
  const mailOptions = {
    from: "ippo7707@gmail.com",
    to: "zacksilaen21@gmail.com",
    subject: "OSM CV Submission",
    text: "Please find the attached XLSX file.",
    attachments: [
      {
        filename: "osmcrew.xlsx",
        content: file,
      },
    ],
  };

  // Send the email
  try {
    const info = await transporter.sendMail(mailOptions);
    console.log("Email sent:", info.response);
  } catch (error) {
    console.error("Error occurred while sending email:", error);
  }
}

module.exports = router;
