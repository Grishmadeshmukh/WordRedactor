const REDACTION_TEXT = "████ REDACTED ████";

const EMAIL = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
const PHONE = /\b(\+?\d{1,3}[\s-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/g;
const SSN = /\b\d{3}-\d{2}-\d{4}\b/g;
const CREDIT_CARD = /\b\d{4}[- ]?\d{4}[- ]?\d{4}[- ]?\d{4}\b/g;
const DOB = /\b(0?[1-9]|1[0-2])[\/\-](0?[1-9]|[12][0-9]|3[01])[\/\-](19|20)\d{2}\b/g;
const MEDICALRECORD = /\bMRN[- ]?\d+\b/gi;
const INSURANCE = /\bINS[- ]?\d+\b/gi;
const EMPLOYEEID = /\bEMP[- ][A-Z0-9-]+\b/gi;
const ADDRESS = /\b\d+\s+[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*\s+(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Boulevard|Blvd|Court|Ct|Way|Circle|Cir|Place|Pl|Terrace)\b/gi;
const SSNLAST4 = /(?:last\s+four|on\s+file|are|is)\s+(?:digits?\s+(?:of\s+the\s+)?(?:social\s+security\s+number\s+)?)?(\d{4})\b/gi;

const PATTERNS = [ EMAIL, PHONE, SSN, CREDIT_CARD, DOB, MEDICALRECORD, 
  INSURANCE, EMPLOYEEID, ADDRESS, SSNLAST4 ];

interface MatchWithCategory {
text: string;
  category: string;
}

//-----------------------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;

  const redactButton = document.getElementById("redactBtn") as HTMLButtonElement;
  const statusDiv = document.getElementById("status");
  if (!redactButton || !statusDiv) return; //not working

  redactButton.addEventListener("click", async () => {
    redactButton.disabled = true;
    statusDiv.textContent = "Redacting document...";
    statusDiv.style.borderLeftColor = "#667eea";

    try {
      //word
      await Word.run(async (context) => {
        if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
          (context.document as any).trackRevisions = true;
        }

        const body = context.document.body;
        body.load("text");
        // console.log("body loaded");
        await context.sync();
        const documentText = body.text;
        // console.log(documentText);

        // match pattrens -------------------------
        PATTERNS.forEach(regex => regex.lastIndex = 0);
        
        const emailMatches = Array.from(documentText.matchAll(EMAIL));
        const phoneMatches = Array.from(documentText.matchAll(PHONE));
        const ssnMatches = Array.from(documentText.matchAll(SSN));
        const creditCardMatches = Array.from(documentText.matchAll(CREDIT_CARD));
        const dobMatches = Array.from(documentText.matchAll(DOB));
        const medicalRecordMatches = Array.from(documentText.matchAll(MEDICALRECORD));
        const insuranceMatches = Array.from(documentText.matchAll(INSURANCE));
        const employeeIdMatches = Array.from(documentText.matchAll(EMPLOYEEID));
        const addressMatches = Array.from(documentText.matchAll(ADDRESS));
        const ssnLast4Matches = Array.from(documentText.matchAll(SSNLAST4)).map(m => m[1] || m[0]);

        const matchesWithCategory: MatchWithCategory[] = [
          ...emailMatches.map(m => ({ text: m[0], category: "Email Addresses" })),
          ...phoneMatches.map(m => ({ text: m[0], category: "Phone Numbers" })),
          ...ssnMatches.map(m => ({ text: m[0], category: "Social Security Numbers" })),
          ...creditCardMatches.map(m => ({ text: m[0], category: "Credit Card Numbers" })),
          ...dobMatches.map(m => ({ text: m[0], category: "Dates of Birth" })),
          ...medicalRecordMatches.map(m => ({ text: m[0], category: "Medical Record Numbers" })),
          ...insuranceMatches.map(m => ({ text: m[0], category: "Insurance Numbers" })),
          ...employeeIdMatches.map(m => ({ text: m[0], category: "Employee IDs" })),
          ...addressMatches.map(m => ({ text: m[0], category: "Addresses" })),
          ...ssnLast4Matches.map(m => ({ text: m, category: "SSN Last 4 Digits" }))
        ];

        // match pattrens -------------------------

        // dedupe 
        const uniqueMatchesMap = new Map<string, string>();
        matchesWithCategory.forEach(({ text, category }) => {
          if (!uniqueMatchesMap.has(text)) { uniqueMatchesMap.set(text, category);
        }});

        const uniqueMatches = Array.from(uniqueMatchesMap.keys());
        let redactionCount = 0;
        const redactedCategories = new Map<string, number>();


// redact -------------------------
        for (const matchText of uniqueMatches) {
          try {
            let ranges = body.search(matchText, { matchCase: false, matchWholeWord: false });
            ranges.load("items");
            await context.sync();

            if (ranges.items.length === 0 && /[- ]/.test(matchText)) {
              const simplifiedText = matchText.replace(/[- ]/g, '');
              ranges = body.search(simplifiedText, { matchCase: false, matchWholeWord: false });
              ranges.load("items");
              await context.sync();
            }

            if (ranges.items.length > 0) {
              const category = uniqueMatchesMap.get(matchText) || "Unknown";
              ranges.items.forEach(range => {
                range.insertText(REDACTION_TEXT, Word.InsertLocation.replace);
                redactionCount++;
                redactedCategories.set(category, (redactedCategories.get(category) || 0) + 1);
              });
              await context.sync();
            }
          } catch (error) {
            console.error(`Error searching for "${matchText}":`, error);
          }
        }

        //  confidential header
        const section = context.document.sections.getFirst();
        const header = section.getHeader("Primary");
        header.insertText("CONFIDENTIAL DOCUMENT\n", Word.InsertLocation.start);
        await context.sync();

        if (redactionCount === 0) {
          statusDiv.textContent = `No sensitive information found. (${matchesWithCategory.length} regex matches)`;
          statusDiv.style.borderLeftColor = "#f59e0b";
        } else {
          const categoriesList = Array.from(redactedCategories.entries())
            .map(([category, count]) => `• ${category}: ${count}`)
            .join("\n");
          statusDiv.textContent = `Redaction complete. ${redactionCount} items redacted:\n\n${categoriesList}`;
          statusDiv.style.borderLeftColor = "#10b981";
        }
        redactButton.disabled = false;
      });
    } catch (err) {
      console.error(err);
      statusDiv.textContent = "Error during redaction";
      statusDiv.style.borderLeftColor = "#ef4444";
      redactButton.disabled = false;
    }
  });
});
