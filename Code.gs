const SPREADSHEET_ID = "1s9LehtM0ZK0SIqBDqg-RPL-5piTjazOkOtfNsgPRsDY";
const DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1512958350906626079/PA7fMaW2TXr01LlHoONMMWSktB5UXi5lpIWtqXy4h5uGHDlRV6_3oHgb3qeAspKyQYht";

const GRAD_MAP = {
  "Grad":     "Graduate",
  "UG":       "Student",
  "Gap-Year": "GapYear",
  "Drop-Out": "DropOut"
};

const NAT_MAP = {
  "Egyptian": "Egyptian",
  "Forigner": "Foreign"
};

const EXP_MAP = {
  "No , I Don't":       "Any",
  "Customer Services":  "CS",
  "Telesales":          "Telesales",
  "Cold Caling":        "CS-Telesales",
  "Recruitmet (only for out 4u offer)": "Any"
};

const FROM_MAP = {
  "Cairo":                    "Cairo",
  "Giza":                     "Giza",
  "Alex":                     "Alex",
  "Other, Can't":             "Cant",
  "Other, can work in Alex":  "Any",
  "Other, can work in Cairo": "Any",
  "Other, can work in Giza":  "Any"
};

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "getSettings") {
    return ContentService.createTextOutput(JSON.stringify(getSettingsData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === "getCompanies") {
    const companies = getFilteredCompanies(
      parseInt(e.parameter.age),
      e.parameter.graduation,
      e.parameter.nationality,
      e.parameter.from,
      e.parameter.experience
    );
    return ContentService.createTextOutput(JSON.stringify(companies))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('FireHire RS Application')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var data;
  if (e.postData && e.postData.contents) {
    data = JSON.parse(e.postData.contents);
  } else {
    data = e.parameter;
  }
  var result = processForm(data);
  return ContentService.createTextOutput(JSON.stringify({result: result}))
    .setMimeType(ContentService.MimeType.JSON);
}

const FIREHIRE_RS_ID = "15Cb7QjzLXsUdicmNur2OnT_DYrw263-bbTSye2vWNbU";

function getSettingsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const settingsSheet = ss.getSheetByName("Settings");
  const data = settingsSheet.getDataRange().getValues();
  let options = { teamLeaderName: [], graduation: [], nationality: [], from: [], experienceInCS: [] };
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) options.teamLeaderName.push(data[i][1].toString().trim());
    if (data[i][2]) options.graduation.push(data[i][2].toString().trim());
    if (data[i][3]) options.nationality.push(data[i][3].toString().trim());
    if (data[i][4]) options.from.push(data[i][4].toString().trim());
    if (data[i][5]) options.experienceInCS.push(data[i][5].toString().trim());
  }

  // ✅ جيب أسماء الريكروترز من شيت FireHire Recruitment Solutions
  try {
    const rsSheet = SpreadsheetApp.openById(FIREHIRE_RS_ID).getSheets()[0];
    const rsData  = rsSheet.getDataRange().getValues();
    const headers = rsData[0].map(h => h.toString().toLowerCase().trim());
    const col     = headers.findIndex(h => h.includes("recruiter name"));
    if (col !== -1) {
      const names = [];
      for (let i = 1; i < rsData.length; i++) {
        const val = rsData[i][col];
        if (val && !names.includes(val.toString().trim())) {
          names.push(val.toString().trim());
        }
      }
      options.recruiterNames = names.sort();
    } else {
      options.recruiterNames = [];
    }
  } catch(e) {
    Logger.log("Recruiter names error: " + e);
    options.recruiterNames = [];
  }

  return options;
}

// ============================================================
// الفلترة الذكية
// ============================================================
function getFilteredCompanies(age, gradRaw, natRaw, fromRaw, expRaw) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Companies");
  if (!sheet) return [];

  const grad = GRAD_MAP[gradRaw] || gradRaw;
  const nat  = NAT_MAP[natRaw]  || natRaw;
  const from = FROM_MAP[fromRaw] || fromRaw;
  const exp  = EXP_MAP[expRaw]  || expRaw;

  const isStudent    = (grad === "Student");
  const isGraduate   = !isStudent;
  const hasExp       = (exp !== "Any");

  const data = sheet.getDataRange().getValues();
  const filtered = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const companyName = row[0].toString().trim();
    const minAge      = parseInt(row[1]);
    const maxAge      = parseInt(row[2]);
    const gradOptions = row[3].toString().trim().split("-");
    const natOptions  = row[4].toString().trim().split("-");
    const fromOptions = row[5].toString().trim().split("-");
    const expOptions  = row[6].toString().trim().split("-");

    // ✅ السن بالظبط
    if (age < minAge || age > maxAge) continue;

    // ✅ التخرج
let gradMatch = false;
if (gradOptions.includes("Any")) {
  gradMatch = true;
} else if (grad === "Graduate") {
  gradMatch = gradOptions.includes("Graduate");
} else if (grad === "Student") {
  gradMatch = gradOptions.includes("Student");
} else if (grad === "GapYear" || grad === "DropOut") {
  gradMatch = !gradOptions.includes("Graduate") || gradOptions.includes("Any");
}

    // ✅ الجنسية
    let natMatch = false;
    if (natOptions.includes("Any")) {
      natMatch = true;
    } else if (nat === "Egyptian") {
      natMatch = natOptions.join("-").includes("Egyptian");
    } else if (nat === "Foreign") {
      natMatch = natOptions.join("-").includes("Foreign");
    }
    if (!natMatch) continue;

    // ✅ المحافظة
    let fromMatch = false;
    if (fromOptions.includes("Any") || from === "Any") {
      fromMatch = true;
    } else {
      fromMatch = fromOptions.includes(from);
    }
    if (!fromMatch) continue;

    // ✅ الخبرة
    // مفيش خبرة → بس Any
    // عنده CS → Any + CS + CS-Telesales
    // عنده Telesales → Any + Telesales + CS-Telesales
    // عنده الاتنين → كل حاجة
    let expMatch = false;
    if (expOptions.includes("Any")) {
      expMatch = true;
    } else if (!hasExp) {
      expMatch = false;
    } else if (exp === "CS") {
      expMatch = expOptions.includes("CS") || expOptions.includes("CS-Telesales");
    } else if (exp === "Telesales") {
      expMatch = expOptions.includes("Telesales") || expOptions.includes("CS-Telesales");
    } else if (exp === "CS-Telesales") {
      expMatch = true;
    }
    if (!expMatch) continue;

    filtered.push(companyName);
  }

  return filtered;
}

// ============================================================
// Discord Notification
// ============================================================
function sendDiscordNotification(formData) {
  try {
    const now = new Date();
    const cairoTime = Utilities.formatDate(now, "Africa/Cairo", "dd/MM/yyyy hh:mm aa");
    const payload = {
      embeds: [{
        title: "🔥 New Application Submitted!",
        color: 0xFF416C,
        thumbnail: { url: "https://i.ibb.co/5Xh4xx8Y/IMG-20251012-WA0004.jpg" },
        fields: [
          { name: "👤 Full Name",        value: formData.fullName       || "N/A", inline: true  },
          { name: "🏢 Company",          value: formData.companyName    || "N/A", inline: true  },
          { name: "\u200B",              value: "\u200B",                         inline: false },
          { name: "👔 Recruiter",        value: formData.recruiter      || "N/A", inline: true  },
          { name: "👑 Team Leader",      value: formData.teamLeaderName || "N/A", inline: true  },
          { name: "\u200B",              value: "\u200B",                         inline: false },
          { name: "📞 Phone",            value: formData.phone          || "N/A", inline: true  },
          { name: "📧 Email",            value: formData.email          || "N/A", inline: true  },
          { name: "\u200B",              value: "\u200B",                         inline: false },
          { name: "🌍 Nationality",      value: formData.nationality    || "N/A", inline: true  },
          { name: "🎓 Graduation",       value: formData.graduation     || "N/A", inline: true  },
          { name: "💼 Experience in CS", value: formData.experienceInCS || "N/A", inline: true  },
          { name: "🎤 Vocaroo Link",     value: formData.vocaroo        || "N/A", inline: false }
        ],
        footer: { text: "FireHire RS | Form Submission • " + cairoTime }
      }]
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: false  // ← غيّرنا ده عشان نشوف الأخطاء
    };

    // محاولة مع Retry تلاتين تاني لو فشل
    let attempts = 0;
    let success = false;

    while (attempts < 3 && !success) {
      try {
        const response = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, options);
        const code = response.getResponseCode();

        Logger.log("Discord response code: " + code);

        if (code === 204 || code === 200) {
          success = true;
          Logger.log("Discord notification sent successfully ✅");
        } else if (code === 429) {
          // Rate limit - استنى
          Logger.log("Rate limited by Discord, waiting...");
          Utilities.sleep(2000);
        } else {
          Logger.log("Discord error body: " + response.getContentText());
          break;
        }
      } catch (innerErr) {
        Logger.log("Attempt " + (attempts + 1) + " failed: " + innerErr.toString());
        Utilities.sleep(1000);
      }
      attempts++;
    }

    if (!success) {
      Logger.log("⚠️ Failed to send Discord notification after " + attempts + " attempts");
    }

  } catch (err) {
    Logger.log("Discord fatal error: " + err.toString());
  }
}

// ============================================================
// Process Form
// ============================================================
function processForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("The Validation");
    if (!sheet) return "Error: Sheet not found!";

    const colAValues = sheet.getRange("A:A").getValues();
    let targetRow = 1;
    for (let i = 0; i < colAValues.length; i++) {
      if (!colAValues[i][0]) { targetRow = i + 1; break; }
    }
    if (targetRow === 1 && colAValues[0][0]) targetRow = colAValues.length + 1;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const findCol = (k) => {
      const idx = headers.findIndex(h => h.toString().toLowerCase().includes(k.toLowerCase()));
      return idx !== -1 ? idx + 1 : null;
    };

    const mapping = [
      { key: "Timestamp",    val: new Date() },
      { key: "Full Name",    val: formData.fullName },
      { key: "Phone Number", val: formData.phone },
      { key: "Email",        val: formData.email },
      { key: "Nationality",  val: formData.nationality },
      { key: "Age",          val: formData.age },
      { key: "Graduation",   val: formData.graduation },
      { key: "From",         val: formData.from },
      { key: "Experience",   val: formData.experienceInCS },
      { key: "Company",      val: formData.companyName },
      { key: "Vocaroo",      val: formData.vocaroo },
      { key: "Recruiter",    val: formData.recruiter },
      { key: "Team Leader",  val: formData.teamLeaderName }
    ];

    mapping.forEach(item => {
      const col = findCol(item.key);
      if (col) sheet.getRange(targetRow, col).setValue(item.val);
    });

    sendDiscordNotification(formData);
    return "SUCCESS_CONFIRMED";
  } catch (e) {
    return "Error: " + e.toString();
  }
}
