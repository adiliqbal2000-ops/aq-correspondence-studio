import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  BookOpen,
  FileText,
  Plus,
  Trash2,
  Copy,
  Check,
  Loader2,
  Send,
  X,
  Search,
  ChevronDown,
  AlertCircle,
  Settings as SettingsIcon,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Design tokens (inline, since only core Tailwind utilities are available)
// ---------------------------------------------------------------------------
const ink = "#1B2A44";      // primary navy
const inkSoft = "#33455F";
const parchment = "#F7F2E7"; // paper background
const parchmentDark = "#EFE7D6";
const charcoal = "#2A2823";
const brass = "#A9803F";
const brassLight = "#C9A15E";
const slate = "#5B6472";
const slateLight = "#8B93A0";
const maroon = "#7B3131";
const line = "#D8CDB2";

const fontImport = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Tinos:ital,wght@0,400;0,700;1,400&family=JetBrains+Mono:wght@400;500&display=swap');
`;

// ---------------------------------------------------------------------------
// AQ style profile — distilled from the letter corpus, used as system context
// ---------------------------------------------------------------------------
const AQ_STYLE_PROFILE = `
You are drafting a business letter in the authentic voice of Abdul Quddoos (AQ), Chief Executive of Al-Tariq Constructors (Pvt.) Ltd, Karachi, Pakistan — a construction/EPC contractor working on power, water, and infrastructure projects. Base the draft on the following observed style profile, distilled from over 100 of his real letters spanning 2014-2026:

STRUCTURE
- Opens by re-establishing the paper trail: "Kindly refer to...", "This refers to...", "In continuation of our letter/meeting on...", "Please refer to our above letters...". Reference letters are often listed as a dated bullet/numbered list when there is a long history.
- For multi-point letters, uses numbered paragraphs (1, 2, 3...) or lettered/numbered itemized lists for amounts, dates, or claims.
- Recites the factual chronology (contract dates, LOA, amounts, letter references) before making the ask — precise about clause numbers, letter reference numbers, and dates.
- For long-running disputes, explicitly acknowledges the passage of time ("despite the passage of nearly six years...") and references his own personal effort ("I personally approached...", "I, tried several times to meet in person...", "the undersigned (CEO-ATL)...") — this personal involvement is a deliberate rhetorical device showing good faith.
- Closes with a clear, specific, quantified ask (amounts in PKR where relevant) and often a graduated consequence ("failing which we will be constrained to seek all available remedies", "to avoid any further complications / potential litigation").
- Sign-off is consistently: "Thanking you and looking forward to your early/kind response." then "Yours faithfully," / "Yours Faithfully," (occasionally "Regards" or "Best regards," for shorter/routine letters), then "For: Al-Tariq Constructors (Pvt.) Ltd." then his name "Abdul Quddoos" then title "Chief Executive" (italicized in original documents).
- CC lines are used to build institutional pressure (widely circulated to directors, engineers, project managers).

TONE — CONTEXT DEPENDENT (important: calibrate to the situation)
- Routine / transactional letters (LC amendments, submissions, routine follow-ups): brief, businesslike, almost no editorializing, sometimes just a table of changes plus "Yours truly."
- Cooperative / early-stage requests: warm but formal — "we would appreciate", "your kind consideration and action would highly be appreciated", "your good office", assurances of continued cooperation and relationship value ("win-win position", "excellent working relations for a long time").
- Firm / escalation letters (payments long overdue, unjust deductions, contractual disputes): direct and unhedged — states the other party's position is wrong plainly ("factually incorrect and contractually invalid", "not maintainable"), rebuilds the record as proof, may note financial hardship being caused, and reserves rights (interest, arbitration, "all available remedies") — but never resorts to personal insult; criticism stays procedural/behavioral ("no positive action has been taken", "not in the spirit of contract").
- Even in firm letters, AQ typically closes on a note of hoping for an amicable / fair resolution rather than pure threat — the firmness and the relationship-preserving close coexist.

RECURRING PHRASES
"Kindly refer to...", "In continuation of our...", "Please be informed...", "We would appreciate if your goodself...", "your good office", "kindly requested", "your kind consideration and action would highly be appreciated", "failing which...", "we reserve our right to...", "win-win position/situation", "based on fairness and honesty", "Thanking you and looking forward to your early/kind response."

FORMATTING CONVENTIONS
- Recipient block: Title/Name, Designation, Company, Address — often with a "KIND ATTN:" line naming a specific contact.
- "PROJECT:" and "SUBJECT:" lines in bold/caps before the salutation, often with a "Reference:" line citing prior correspondence.
- Salutation is "Dear Sir," (rarely anything else).
- Body paragraphs are justified, moderate length (3-6 sentences), formal register throughout — contractions are never used.

Draft the requested letter matching this profile as closely as possible, calibrating tone to the situation described. Do not invent facts, amounts, or dates beyond what is provided — leave a bracketed placeholder like [amount] or [date] only if something essential is missing and cannot be reasonably inferred.
`.trim();

// ---------------------------------------------------------------------------
// Directory config — the "dropdown with add-other" reference lists
// Seeded from the real recipient/subject/project data found across AQ's
// archived correspondence (JPCL, CPGCL, CAA/IIAP, Adamjee, KSB, Bank Al-Habib).
// ---------------------------------------------------------------------------
const DIRECTORY_FIELDS = [
  {
    key: "refPrefix",
    seed: ["ATL/206.1/26-", "ATL/222.7/26-", "ATL-GPL/222.7/26/", "ATL/228.1/26-", "ATL/236.1/26-"],
  },
  {
    key: "addressee",
    seed: [
      "Chief Executive Officer (CEO)",
      "The Chief Engineer / Project Director",
      "Project Director",
      "Project Manager/ER",
      "Joint Senior Manager",
      "Managing Director",
      "Director General (HR & Admn)",
      "Branch Manager",
    ],
  },
  {
    key: "recipientCompany",
    seed: [
      "Jamshoro Power Company Limited (JPCL)",
      "Central Power Generation Company Limited (CPGCL)",
      "Adamjee Insurance Company Limited",
      "Pakistan Civil Aviation Authority (CAA)",
      "KSB Pumps Company Limited",
      "Bank Al-Habib Limited",
      "Karachi Water & Sewerage Board (KW&SB)",
    ],
  },
  {
    key: "recipientAddress",
    seed: [
      "Mohra Jabbal, Dadu Road, Jamshoro",
      "747 MW CCPP Guddu, Genco II, TPS Guddu",
      "Islamabad International Airport (IIAP), Islamabad, Pakistan",
      "16/2, Sir Agha Khan Road, Lahore",
      "New Challi Branch, Haji Adam Chamber, Altaf Hussain Road, Karachi, Pakistan",
      "KW&SB 100 MGD Pump Station, Dhabeji",
    ],
  },
  { key: "kindAttn", seed: [] },
  {
    key: "projectTitle",
    seed: [
      "Construction of Hazardous Solid Waste Disposal Facility and Associated Civil Works (JPCL)",
      "Design, Supply, Installation, Testing & Commissioning of Natural Gas Booster Compressor Station, CPGCL Guddu",
      "Package 6: Hydrant Refueling System for Islamabad International Airport (IIAP)",
      "Construction of New 100 MGD Pump House at Dhabeji & Clifton Pumping Station Project",
    ],
  },
  {
    key: "subject",
    seed: [
      "Overdue Payment",
      "Outstanding Payment",
      "Request for Extension of Time",
      "Unsubstantiated Deductions from Payment Certificate",
      "Encashment of Performance Security / Bank Guarantee",
      "Extension of Performance Guarantee",
      "Variation Order / Prolongation Claim",
      "Request for Idle Time Cost of Man-Power, Equipment, Tool & Plant",
    ],
  },
];

// ---------------------------------------------------------------------------
// Starter letters — verified, clean examples pulled directly from AQ's
// archived correspondence, spanning routine and escalation registers.
// ---------------------------------------------------------------------------
const STARTER_LETTERS = [
  {
    id: "starter-jpcl-2025",
    title: "ATL-GPL/222.7/25/565 — JPCL Overdue Payment",
    client: "Jamshoro Power Company Limited (JPCL)",
    project: "Construction of Hazardous Solid Waste Disposal Facility and Associated Civil Works Project",
    date: "18.07.2025",
    tags: "payment, escalation, variation order",
    content: `Chief Executive Officer (CEO),
Jamshoro Power Company Limited (JPCL)
Mohra Jabbal, Dadu Road, Jamshoro
Phone: 022-9213706

Subject: CONSTRUCTION OF HAZARDOUS SOLID WASTE DISPOSAL FACILITY AND ASSOCIATED CIVIL WORKS PROJECT FOR JAMSHORO POWER COMPANY LTD (BID NO. ADB-L30900-PACKICBJPGP001-3)

Reference: Overdue payment regarding increased quantities of Hazardous Solid Waste.

Dear Sir,

1. Subject project (Works) was awarded to ATL-JPL Consortium (Contractor) pursuant to Letter of Acceptance dated 11.11.2017 (LOA). Contract Agreement was signed between Jamshoro Power Company Limited (JPCL) (Employer) and ATL-JPL on 17.03.2018. The Works commenced on 11.07.2018 to be completed within 180 calendar days i.e. on 06.10.2019. There were delays during the execution of works (not attributable to the Contractor). Therefore, the Engineer (Mott MacDonald Limited - MMP) extended the Time for Completion upto 25.03.2020. The project was successfully completed by the Contractor in terms of the contractual requirements and the same was taken over by the Employer.

2. During execution phase, quantities of Hazardous Solid Waste were increased substantially from 2000 CUM (originally prescribed quantities) to 4800 CUM (increased and additional quantities) by the Engineer in consultation and consent of the Employer. In the circumstances, Contractor submitted Variation Order (VO) on 19.06.2019 for an amount of Rs.82,281,170/- which was checked, verified and recommended by "The Engineer" for an amount of Rs.58,780,647/-. The Final Bill (statement) was submitted by the Contractor in due course.

3. The VO recommended by "the Engineer" and funds for the same amount has been approved by the Asian Development Bank (Financer of the Project) on 31.05.2022 for payment. However, the amount in relation to the increased executed quantities of Hazardous Material is still overdue for payment despite lapse of sufficient period. We refer to Engineer's letter Ref. No. 334014-MML-JPCL-L-00085 dated 23.11.2021 through which the Engineer asked the Contractor to secure approval of the pending VO directly from the Employer. To which, we have already clarified many times that under the provisions of the Contract, the Contractor has nothing to do with the approval of the Employer and once the VO is checked, verified and recommended, the Employer is bound to issue its approval as additional quantities at site regarding the increased quantities of Hazardous Material have already been carried out by us under specific instructions of the Engineer (competent authority).

4. Please note that VO was duly initiated and recommended by the Engineer vide letter Ref. No. 334014-MML-JPCL-L-01393 dated 09.07.2019. The Employer never disputed or repudiated entitlement of the amount recommended and certified by "the Engineer" pursuant to VO. However, has failed to pay and release the same despite our repeated requests as mentioned a few of the following correspondence: -

(i) Contractor's Letter Ref. No. ATL-GPL/222.7/24/482 dated 07.10.2024.
(ii) Contractor's Letter Ref. No. ATL-GPL/222.7/24/315 dated 15.07.2024.
(iii) Contractor's Letter Ref. No. ATL-GPL/222.7/23/637 dated 28.12.2023.
(iv) Contractor's Letter Ref. No. ATL-GPL/222.7/22/437 dated 13.12.2022.
(v) Contractor's Letter Ref. No. ATL-GPL/222.7/21-281 dated 04.10.2021.

5. Please note that in terms of relevant Clauses in the Contract, if the Employer fails to pay the certified/recommended amount of Rs.58,780,467/- to the Contractor in specified time as mentioned in the Contract Agreement, the Contractor becomes entitled to receive Financing Charges compounded monthly on the amount unpaid during the period of delay. Accordingly, Financing Charges calculated upto 30 June, 2025 comes to Rs.58,453,730/-. We have been repeatedly led to believe through meetings and correspondence that the matter will be resolved amicably, however, we feel that no reasonable effort is made till date by the Employer.

In view of this matter, you are kindly requested to arrange the following payments at your earliest failing which we will be constrained to seek all the available remedies before the appropriate forum at your risk and cost.

(i) Verified/Recommended amount by the Engineer — Rs. 58,780,647/-
(ii) Accrued financial charges (delayed Compensation) calculated on average base up to 30.06.2025 (attached as appendix-A) — Rs. 58,453,730/-
TOTAL:- Rs. 117,234,377/-

Yours faithfully,
For: Al-Tariq Constructor (Pvt) Ltd.,

Abdul Quddoos
Chief Executive Officer`,
  },
  {
    id: "starter-caa-2023",
    title: "ATL/206.1/23-524 — CAA/IIAP Payment Release",
    client: "Pakistan Civil Aviation Authority (CAA)",
    project: "Package 6: Hydrant Refueling System for Islamabad International Airport Project",
    date: "08.11.2023",
    tags: "payment, arbitration, firm",
    content: `The Project Director,
Pakistan Civil Aviation Authority
Islamabad International Airport (IIAP)
Islamabad, Pakistan

PROJECT: PACKAGE 6: HYDRANT REFUELING SYSTEM FOR ISLAMABAD INTERNATIONAL AIRPORT PROJECT.

Dear Sir,

Kindly, refer to your Letter No: PD(IIAP)/6279/11/E&M /302 dated 10th July 2023 in response to our various letters, wherein we have requested you to release our payment as per the verified and passed invoice from the consultant, your response as referred above is not in line. You are therefore finally requested to release our payment as per the passed and verified invoice from Engineer along with the prolongation claim or otherwise, we reserve the right to invoke the contract clause of arbitration.

We hope that your good office will strictly follow the Contract clauses to safeguard the interest of the Contractor based on justice and fairness and release all the payment which is the right of the Contractor.

We again attached herewith the summary of payment for your reference and necessary action.

Thanking you and looking forward to your early response as our payment is on hold without any reason, we further reserve the right to claim the interest on our payment being held since we are facing a lot of financial hardship.

Regards
For: Al-Tariq Constructors (Pvt.) Ltd

Abdul Quddoos
Chief Executive

Cc:
Kind information: Director General – CAA. (Address: Jinnah Terminal-1 Building, Karachi.)

Encl:
Final Account Calculation Sheet (Annex-A)
Copy of Passed Invoice by the Engineer`,
  },
  {
    id: "starter-ksb-2025",
    title: "ATL/236.1/25- — KSB Outstanding Payment",
    client: "KSB Pumps Company Limited",
    project: "Construction of New 100 MGD Pump House at Dhabeji & Clifton Pumping Station Project",
    date: "03.11.2025",
    tags: "payment, personal effort, long-running dispute",
    content: `MR. IMRAN GHANI
Managing Director,
KSB Pumps Company Limited
16/2, Sir Agha Khan Road,
Lahore.

PROJECT: CONSTRUCTION OF NEW 100 MGD PUMP HOUSE AT DHABEJI & CLIFTON PUMPING STATION PROJECT
SUBJECT: OUTSTANDING PAYMENT

Dear Sir,

Please be inform once again that our outstanding payment as following is still to be unpaid or decided:

1. Construction of New 100 MGD Pump House: (Equipped with Mechanical & Electrical) Pumping Machineries at Dhabeji Pumping Station. Amounting to Rs. 72,790,596/- (attached)

2. Construction of Clifton Pumping station (sewage): Adjacent to CID Centre, Karachi. Amounting to Rs. 33,074,319/- (attached)

It is a matter of record that I, tried several times to meet in person with your good self but failed. Please refer to your email dated 15th September 2024, wherein you suggested informing us of a convenient time, and to your second email dated 18th September 2024, in which you suggested meeting Mr. Imran Malik. I have held several meetings with Mr. Imran Malik; he called Mr. Farooq and advised meeting at the Karachi ATL office to finalize the payment issues together with the pending payments toward the Karachi Water & Sewerage Board, especially for above referred projects. However, despite these efforts, no positive action has been taken, although your good office had already addressed the Managing Director, KW&SB, to release the payment vide letter No. CSZ-084/CPS/2013-801 dated 15th April 2025.

Keeping in view all the above facts, nobody met at the ATL Karachi office, and no meeting has been held by the KSB Lahore office. As per your earlier suggestion, we trust that your good self will kindly adhere to your commitments and take the necessary steps to resolve the outstanding issues to avoid any further complications. The matter has been pending for a long time and requires your personal attention.

Thanking you and looking forward to your early response.

Yours Faithfully,

Abdul Quddoos
Chief Executive

Cc:
Mr. Imran Malik -- Director (KSB)

Encl:
Email dated 15th September 2024
Email dated 18th September 2024
KSB Letter No. CSZ-084/CPS/2013-801 dated 15th April 2025.
ATL Letter No. ATL/236.1/23-526 dated 08th November 2023.`,
  },
  {
    id: "starter-bankalhabib-2016",
    title: "ATL/228.1/ — Bank Al-Habib LC Amendment",
    client: "Bank Al-Habib Limited",
    project: "",
    date: "08-12-2016",
    tags: "routine, transactional, letter of credit",
    content: `Branch Manager,
Bank Al Habib Limited,
New Challi Branch,
Haji Adam Chamber,
Altaf Hussain Road, Karachi.
Phone: 021-32410526

SUBJECT: Amendment in Letter of Credit
1015LC44211/2016 dated 07-11-2016

Dear Sir,

This has reference to the above letter of credit, our supplier M/s. Focal Solution China has requested to amend the following clause as under:

46A1) 'BENEFICIARY'S .........MERCHANDISE ARE OF TAIWAN ORIGIN MENTIONING ...........'

To

46A1) 'BENEFICIARY'S ..................MERCHANDISE ARE OF CHINA ORIGIN MENTIONING ...........'

You are requested to please make the amendment and inform us accordingly.

Yours faithfully,
For: Al-Tariq Constructors (Pvt.) Ltd.,

Abdul Quddoos
Chief Executive`,
  },
];

function dedupeMerge(existing, incoming) {
  const result = [...existing];
  const lower = new Set(existing.map((v) => v.toLowerCase()));
  for (const v of incoming) {
    const trimmed = (v || "").trim();
    if (trimmed && !lower.has(trimmed.toLowerCase())) {
      result.push(trimmed);
      lower.add(trimmed.toLowerCase());
    }
  }
  return result;
}

function extractRefPrefixes(text) {
  if (!text) return [];
  const matches = text.match(/ATL(?:-GPL)?\/\d{2,3}\.\d\/\d{2}[-/]/g) || [];
  return [...new Set(matches)];
}

// ---------------------------------------------------------------------------
// Storage helpers
// ---------------------------------------------------------------------------
// localStorage-backed persistence (standalone version — no Claude.ai runtime needed)
const LS_LIBRARY_KEY = "aqcs:library";
const LS_DIRECTORY_PREFIX = "aqcs:directory:";
const LS_SETTINGS_KEY = "aqcs:settings";

async function loadLibrary() {
  try {
    const raw = localStorage.getItem(LS_LIBRARY_KEY);
    if (!raw) return [];
    const obj = JSON.parse(raw);
    const entries = Object.values(obj);
    entries.sort((a, b) => (b.addedAt || 0) - (a.addedAt || 0));
    return entries;
  } catch (e) {
    return [];
  }
}

async function saveLibraryEntry(entry) {
  try {
    const raw = localStorage.getItem(LS_LIBRARY_KEY);
    const obj = raw ? JSON.parse(raw) : {};
    obj[entry.id] = entry;
    localStorage.setItem(LS_LIBRARY_KEY, JSON.stringify(obj));
  } catch (e) {
    // ignore
  }
}

async function deleteLibraryEntry(id) {
  try {
    const raw = localStorage.getItem(LS_LIBRARY_KEY);
    const obj = raw ? JSON.parse(raw) : {};
    delete obj[id];
    localStorage.setItem(LS_LIBRARY_KEY, JSON.stringify(obj));
  } catch (e) {
    // ignore
  }
}

async function loadDirectory(key) {
  try {
    const raw = localStorage.getItem(LS_DIRECTORY_PREFIX + key);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

async function saveDirectory(key, arr) {
  try {
    localStorage.setItem(LS_DIRECTORY_PREFIX + key, JSON.stringify(arr));
  } catch (e) {
    // ignore
  }
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_SETTINGS_KEY);
    return raw ? JSON.parse(raw) : { apiKey: "", model: "claude-sonnet-5" };
  } catch (e) {
    return { apiKey: "", model: "claude-sonnet-5" };
  }
}

function saveSettings(settings) {
  try {
    localStorage.setItem(LS_SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) {
    // ignore
  }
}

// ---------------------------------------------------------------------------
// Combobox — text input + dropdown of saved options + "add other"
// ---------------------------------------------------------------------------
function Combobox({ label, value, onChange, options, placeholder, onAddOption }) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef(null);

  useEffect(() => {
    function handleClickOutside(e) {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const query = value || "";
  const filtered = (options || []).filter((o) => o.toLowerCase().includes(query.toLowerCase()));
  const exactMatch = (options || []).some((o) => o.toLowerCase() === query.trim().toLowerCase());

  function selectOption(opt) {
    onChange(opt);
    setOpen(false);
  }

  function handleAddCustom() {
    const val = query.trim();
    if (!val) return;
    if (onAddOption) onAddOption(val);
    setOpen(false);
  }

  return (
    <div className="relative" ref={containerRef}>
      <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>{label}</label>
      <div className="relative mt-1">
        <input
          value={query}
          onChange={(e) => {
            onChange(e.target.value);
            setOpen(true);
          }}
          onFocus={() => setOpen(true)}
          placeholder={placeholder}
          className="w-full rounded-md border py-2 pl-3 pr-8 text-sm outline-none"
          style={{ borderColor: line }}
        />
        <button
          type="button"
          onClick={() => setOpen((v) => !v)}
          className="absolute right-2 top-1/2 -translate-y-1/2"
          style={{ color: slateLight }}
        >
          <ChevronDown size={14} style={{ transform: open ? "rotate(180deg)" : "none" }} />
        </button>
      </div>
      {open && (
        <div
          className="absolute z-20 mt-1 max-h-48 w-full overflow-y-auto rounded-md border bg-white shadow-lg"
          style={{ borderColor: line }}
        >
          {filtered.length === 0 && (
            <div className="px-3 py-2 text-xs" style={{ color: slateLight }}>
              {options && options.length > 0 ? "No matches" : "No saved options yet"}
            </div>
          )}
          {filtered.map((opt) => (
            <div
              key={opt}
              onClick={() => selectOption(opt)}
              className="cursor-pointer px-3 py-1.5 text-xs hover:bg-gray-50"
              style={{ color: charcoal }}
            >
              {opt}
            </div>
          ))}
          {query.trim() && !exactMatch && (
            <div
              onClick={handleAddCustom}
              className="cursor-pointer border-t px-3 py-1.5 text-xs font-medium hover:bg-gray-50"
              style={{ borderColor: line, color: brass }}
            >
              + Add "{query.trim()}" as new option
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function Field({ label, value, onChange, placeholder }) {
  return (
    <div>
      <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>{label}</label>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
        style={{ borderColor: line }}
      />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Letterhead components (static — the identity element of the app)
// ---------------------------------------------------------------------------
function LetterheadATL() {
  return (
    <div>
      <div className="flex items-start justify-between">
        <div className="flex items-start gap-3">
          <div
            className="flex h-12 w-12 flex-shrink-0 items-center justify-center rounded-sm border-2 font-bold"
            style={{ borderColor: brass, color: ink, fontFamily: "Inter" }}
          >
            <span style={{ fontSize: "13px", letterSpacing: "0.5px" }}>ATL</span>
          </div>
          <div>
            <div
              className="font-bold"
              style={{ color: ink, fontFamily: "Inter", fontSize: "15px", letterSpacing: "0.5px" }}
            >
              AL-TARIQ CONSTRUCTORS (PVT.) LTD.
            </div>
            <div style={{ color: slate, fontFamily: "Inter", fontSize: "10.5px", lineHeight: 1.5, marginTop: "2px" }}>
              Suite: 1301-1302, 13th Floor, Uni Centre, I.I Chundrigar Road, Karachi - 74000, Pakistan.
              <br />
              Tel: 0092-21-3242-7800, 3242-7820, Fax: 0092-21-3242-7784
              <br />
              URL: www.atlpk.com &nbsp;|&nbsp; E-mail: atl@atlpk.com
            </div>
          </div>
        </div>
        <div className="text-right" style={{ fontFamily: "Inter", fontSize: "9px", color: slateLight }}>
          <div className="rounded-sm border px-2 py-1" style={{ borderColor: line }}>
            ISO 9001 &middot; ISO 14001 &middot; OHSAS 18001
          </div>
        </div>
      </div>
      <div className="mt-3 flex items-center justify-between">
        <div className="h-px flex-1" style={{ backgroundColor: brass }} />
      </div>
      <div className="mt-1 text-right italic" style={{ color: brass, fontFamily: "Tinos", fontSize: "11px" }}>
        Total Commitment — Our Core Value
      </div>
    </div>
  );
}

function LetterheadConsortium() {
  return (
    <div>
      <div className="flex items-center justify-between">
        <div
          className="flex h-11 w-11 flex-shrink-0 items-center justify-center rounded-sm border-2 font-bold"
          style={{ borderColor: brass, color: ink }}
        >
          <span style={{ fontSize: "12px" }}>ATL</span>
        </div>
        <div
          className="font-bold"
          style={{ color: ink, fontFamily: "Inter", fontSize: "17px", letterSpacing: "1px" }}
        >
          ATL &ndash; GPL CONSORTIUM
        </div>
        <div className="text-right" style={{ fontFamily: "Inter", fontSize: "9px", color: slateLight }}>
          <div className="rounded-sm border px-2 py-1" style={{ borderColor: line }}>
            ISO 9001 &middot; ISO 14001 &middot; OHSAS 18001
          </div>
        </div>
      </div>
      <div className="mt-3 h-px" style={{ backgroundColor: brass }} />
    </div>
  );
}

function LetterFooterConsortium() {
  return (
    <div className="mt-8 border-t pt-2 text-right" style={{ borderColor: line, fontFamily: "Inter", fontSize: "9px", color: slateLight }}>
      <div style={{ fontWeight: 600, color: slate }}>ATL - GPL CONSORTIUM.</div>
      <div>Suit No. 1301, 1302, 13th Floor, Uni Center, I.I Chundrigar Road, Karachi &ndash; Pakistan</div>
      <div>Phone: 021-32427800-3 &nbsp; Email: atl@atlpk.com, info@atlpk.com</div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main App
// ---------------------------------------------------------------------------
export default function App() {
  const [tab, setTab] = useState("draft");
  const [library, setLibrary] = useState([]);
  const [libraryLoading, setLibraryLoading] = useState(true);
  const [showSettings, setShowSettings] = useState(false);
  const [hasApiKey, setHasApiKey] = useState(!!loadSettings().apiKey);
  const [directories, setDirectories] = useState(
    Object.fromEntries(DIRECTORY_FIELDS.map((f) => [f.key, f.seed]))
  );

  // initial load: library + persisted directories
  useEffect(() => {
    (async () => {
      setLibraryLoading(true);
      const entries = await loadLibrary();
      setLibrary(entries);

      const loaded = {};
      for (const f of DIRECTORY_FIELDS) {
        const persisted = await loadDirectory(f.key);
        loaded[f.key] = dedupeMerge(f.seed, persisted);
      }
      setDirectories(loaded);
      setLibraryLoading(false);
    })();
  }, []);

  // auto-seed directories from library content whenever it changes
  useEffect(() => {
    if (library.length === 0) return;
    setDirectories((prev) => {
      const next = { ...prev };
      const companies = library.map((e) => e.client).filter(Boolean);
      const projects = library.map((e) => e.project).filter(Boolean);
      const refPrefixes = library.flatMap((e) => extractRefPrefixes(`${e.title || ""} ${e.content || ""}`));
      next.recipientCompany = dedupeMerge(next.recipientCompany || [], companies);
      next.projectTitle = dedupeMerge(next.projectTitle || [], projects);
      next.refPrefix = dedupeMerge(next.refPrefix || [], refPrefixes);
      return next;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [library]);

  const refreshLibrary = useCallback(async () => {
    const entries = await loadLibrary();
    setLibrary(entries);
  }, []);

  const addDirectoryValue = useCallback((key, value) => {
    setDirectories((prev) => {
      const existing = prev[key] || [];
      if (existing.some((v) => v.toLowerCase() === value.toLowerCase())) return prev;
      const updated = [...existing, value];
      saveDirectory(key, updated);
      return { ...prev, [key]: updated };
    });
  }, []);

  return (
    <div style={{ backgroundColor: "#EDE8DC", minHeight: "100vh", fontFamily: "Inter" }}>
      <style>{fontImport}</style>

      {/* Top bar */}
      <div style={{ backgroundColor: ink }} className="sticky top-0 z-10 shadow-md">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-2.5">
            <div
              className="flex h-8 w-8 items-center justify-center rounded-sm border font-bold text-white"
              style={{ borderColor: brassLight, fontSize: "11px" }}
            >
              AQ
            </div>
            <div>
              <div className="text-white" style={{ fontSize: "15px", fontWeight: 600, letterSpacing: "0.2px" }}>
                Correspondence Studio
              </div>
              <div style={{ color: "#9FADC4", fontSize: "10.5px" }}>Abdul Quddoos &middot; voice-matched drafting</div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex gap-1 rounded-md p-1" style={{ backgroundColor: "#152036" }}>
              <button
                onClick={() => setTab("draft")}
                className="flex items-center gap-1.5 rounded px-3 py-1.5 text-sm font-medium transition"
                style={{
                  backgroundColor: tab === "draft" ? brass : "transparent",
                  color: tab === "draft" ? ink : "#C7CEDB",
                }}
              >
                <FileText size={14} /> New Draft
              </button>
              <button
                onClick={() => setTab("library")}
                className="flex items-center gap-1.5 rounded px-3 py-1.5 text-sm font-medium transition"
                style={{
                  backgroundColor: tab === "library" ? brass : "transparent",
                  color: tab === "library" ? ink : "#C7CEDB",
                }}
              >
                <BookOpen size={14} /> Style Library
                <span
                  className="ml-0.5 rounded-full px-1.5 text-[10px]"
                  style={{ backgroundColor: tab === "library" ? "rgba(27,42,68,0.15)" : "#26314A", color: tab === "library" ? ink : "#9FADC4" }}
                >
                  {library.length}
                </span>
              </button>
            </div>
            <button
              onClick={() => setShowSettings(true)}
              className="flex items-center gap-1.5 rounded-md border px-3 py-1.5 text-xs font-medium"
              style={{
                borderColor: hasApiKey ? "#2E4A38" : "#7A3B3B",
                backgroundColor: hasApiKey ? "#1E2E24" : "#2E1E1E",
                color: hasApiKey ? "#8FD4A6" : "#E3A0A0",
              }}
              title={hasApiKey ? "API key set" : "API key needed"}
            >
              <SettingsIcon size={13} />
              {hasApiKey ? "Settings" : "Add API key"}
            </button>
          </div>
        </div>
      </div>

      {showSettings && (
        <SettingsModal
          onClose={() => {
            setShowSettings(false);
            setHasApiKey(!!loadSettings().apiKey);
          }}
        />
      )}

      <div className="mx-auto max-w-6xl px-6 py-8">
        {tab === "draft" ? (
          <DraftTab
            library={library}
            onSavedToLibrary={refreshLibrary}
            directories={directories}
            addDirectoryValue={addDirectoryValue}
          />
        ) : (
          <LibraryTab
            library={library}
            loading={libraryLoading}
            onChange={refreshLibrary}
            directories={directories}
            addDirectoryValue={addDirectoryValue}
          />
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Library Tab
// ---------------------------------------------------------------------------
function LibraryTab({ library, loading, onChange, directories, addDirectoryValue }) {
  const [showAdd, setShowAdd] = useState(false);
  const [query, setQuery] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [importing, setImporting] = useState(false);

  const filtered = library.filter((e) => {
    const q = query.toLowerCase();
    if (!q) return true;
    return (
      (e.title || "").toLowerCase().includes(q) ||
      (e.client || "").toLowerCase().includes(q) ||
      (e.project || "").toLowerCase().includes(q) ||
      (e.tags || "").toLowerCase().includes(q)
    );
  });

  const existingIds = new Set(library.map((e) => e.id));
  const pendingStarters = STARTER_LETTERS.filter((s) => !existingIds.has(s.id));

  async function handleDelete(id) {
    await deleteLibraryEntry(id);
    onChange();
  }

  async function handleImportStarters() {
    setImporting(true);
    try {
      for (const letter of pendingStarters) {
        await saveLibraryEntry({ ...letter, addedAt: Date.now() });
      }
      await onChange();
    } finally {
      setImporting(false);
    }
  }

  return (
    <div>
      <div className="mb-5 flex items-center justify-between">
        <div>
          <h2 style={{ color: ink, fontFamily: "Inter", fontSize: "18px", fontWeight: 700 }}>Style Library</h2>
          <p style={{ color: slate, fontSize: "12.5px", marginTop: "2px" }}>
            Sample letters AQ has written. The more you add, the better new drafts will match his voice.
          </p>
        </div>
        <div className="flex gap-2">
          {pendingStarters.length > 0 && (
            <button
              onClick={handleImportStarters}
              disabled={importing}
              className="flex items-center gap-1.5 rounded-md border px-3.5 py-2 text-sm font-medium shadow-sm"
              style={{ borderColor: brass, color: brass, backgroundColor: "#FBF3E4", opacity: importing ? 0.7 : 1 }}
            >
              {importing ? <Loader2 size={15} className="animate-spin" /> : <BookOpen size={15} />}
              Import {pendingStarters.length} verified letter{pendingStarters.length > 1 ? "s" : ""} from archive
            </button>
          )}
          <button
            onClick={() => setShowAdd(true)}
            className="flex items-center gap-1.5 rounded-md px-3.5 py-2 text-sm font-medium text-white shadow-sm"
            style={{ backgroundColor: ink }}
          >
            <Plus size={15} /> Add letter
          </button>
        </div>
      </div>

      <div className="relative mb-5">
        <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2" style={{ color: slateLight }} />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search by client, project, or tag..."
          className="w-full rounded-md border py-2 pl-9 pr-3 text-sm outline-none"
          style={{ borderColor: line, backgroundColor: "white" }}
        />
      </div>

      {loading ? (
        <div className="flex items-center gap-2 py-12 justify-center" style={{ color: slate }}>
          <Loader2 size={18} className="animate-spin" /> Loading library...
        </div>
      ) : filtered.length === 0 ? (
        <div
          className="rounded-lg border border-dashed py-14 text-center"
          style={{ borderColor: line, color: slate }}
        >
          <BookOpen size={26} className="mx-auto mb-2" style={{ color: slateLight }} />
          <div style={{ fontSize: "13.5px" }}>
            {library.length === 0 ? "No letters saved yet. Add AQ's past letters to build the style profile." : "No letters match your search."}
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
          {filtered.map((entry) => (
            <div
              key={entry.id}
              className="rounded-lg border bg-white p-4 shadow-sm"
              style={{ borderColor: line }}
            >
              <div className="flex items-start justify-between gap-2">
                <div>
                  <div style={{ fontWeight: 600, color: ink, fontSize: "13.5px" }}>{entry.title || "Untitled letter"}</div>
                  <div style={{ color: slate, fontSize: "11.5px", marginTop: "1px" }}>
                    {entry.client}
                    {entry.client && entry.project ? " · " : ""}
                    {entry.project}
                  </div>
                </div>
                <button onClick={() => handleDelete(entry.id)} title="Delete" style={{ color: slateLight }}>
                  <Trash2 size={15} />
                </button>
              </div>
              {entry.date && (
                <div className="mt-1.5 inline-block rounded px-1.5 py-0.5" style={{ fontSize: "10px", color: brass, backgroundColor: "#FBF3E4", fontFamily: "JetBrains Mono" }}>
                  {entry.date}
                </div>
              )}
              <p
                className="mt-2 cursor-pointer"
                style={{
                  color: charcoal,
                  fontSize: "12px",
                  fontFamily: "Tinos",
                  display: "-webkit-box",
                  WebkitLineClamp: expanded === entry.id ? "none" : 3,
                  WebkitBoxOrient: "vertical",
                  overflow: "hidden",
                }}
                onClick={() => setExpanded(expanded === entry.id ? null : entry.id)}
              >
                {entry.content}
              </p>
              <button
                onClick={() => setExpanded(expanded === entry.id ? null : entry.id)}
                style={{ color: brass, fontSize: "11px", marginTop: "4px" }}
              >
                {expanded === entry.id ? "Show less" : "Read more"}
              </button>
            </div>
          ))}
        </div>
      )}

      {showAdd && (
        <AddLetterModal
          onClose={() => setShowAdd(false)}
          onSaved={onChange}
          directories={directories}
          addDirectoryValue={addDirectoryValue}
        />
      )}
    </div>
  );
}

function AddLetterModal({ onClose, onSaved, directories, addDirectoryValue }) {
  const [title, setTitle] = useState("");
  const [client, setClient] = useState("");
  const [project, setProject] = useState("");
  const [date, setDate] = useState("");
  const [tags, setTags] = useState("");
  const [content, setContent] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  async function handleSave() {
    if (!content.trim()) {
      setError("Paste the letter text before saving.");
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const entry = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        title: title.trim() || "Untitled letter",
        client: client.trim(),
        project: project.trim(),
        date: date.trim(),
        tags: tags.trim(),
        content: content.trim(),
        addedAt: Date.now(),
      };
      await saveLibraryEntry(entry);
      await onSaved();
      onClose();
    } catch (e) {
      setError("Couldn't save this letter. Please try again.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-20 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="max-h-[85vh] w-full max-w-xl overflow-y-auto rounded-lg bg-white p-5 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between">
          <h3 style={{ color: ink, fontWeight: 700, fontSize: "15px" }}>Add a letter to the library</h3>
          <button onClick={onClose} style={{ color: slateLight }}>
            <X size={18} />
          </button>
        </div>

        <div className="space-y-3">
          <Field label="Title / reference no." value={title} onChange={setTitle} placeholder="e.g. ATL-GPL/222.7/25/565 — JPCL overdue payment" />
          <div className="grid grid-cols-2 gap-3">
            <Combobox
              label="Client / recipient"
              value={client}
              onChange={setClient}
              options={directories.recipientCompany}
              placeholder="e.g. JPCL"
              onAddOption={(v) => {
                setClient(v);
                addDirectoryValue("recipientCompany", v);
              }}
            />
            <Combobox
              label="Project"
              value={project}
              onChange={setProject}
              options={directories.projectTitle}
              placeholder="e.g. Hazardous Waste Facility"
              onAddOption={(v) => {
                setProject(v);
                addDirectoryValue("projectTitle", v);
              }}
            />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Date" value={date} onChange={setDate} placeholder="e.g. 18.07.2025" />
            <Field label="Tags" value={tags} onChange={setTags} placeholder="e.g. payment, dispute" />
          </div>
          <div>
            <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>Letter text</label>
            <textarea
              value={content}
              onChange={(e) => setContent(e.target.value)}
              rows={10}
              placeholder="Paste the full letter text here..."
              className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
              style={{ borderColor: line, fontFamily: "Tinos" }}
            />
          </div>
          {error && (
            <div className="flex items-center gap-1.5" style={{ color: maroon, fontSize: "12px" }}>
              <AlertCircle size={13} /> {error}
            </div>
          )}
        </div>

        <div className="mt-5 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-md px-3.5 py-2 text-sm" style={{ color: slate }}>
            Cancel
          </button>
          <button
            onClick={handleSave}
            disabled={saving}
            className="flex items-center gap-1.5 rounded-md px-4 py-2 text-sm font-medium text-white"
            style={{ backgroundColor: ink, opacity: saving ? 0.7 : 1 }}
          >
            {saving ? <Loader2 size={14} className="animate-spin" /> : <Check size={14} />}
            Save to library
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Draft Tab
// ---------------------------------------------------------------------------
function DraftTab({ library, onSavedToLibrary, directories, addDirectoryValue }) {
  const today = new Date();
  const defaultDate = `${String(today.getDate()).padStart(2, "0")}.${String(today.getMonth() + 1).padStart(2, "0")}.${today.getFullYear()}`;

  const [letterhead, setLetterhead] = useState("atl");
  const [addressee, setAddressee] = useState("");
  const [recipientCompany, setRecipientCompany] = useState("");
  const [recipientAddress, setRecipientAddress] = useState("");
  const [attention, setAttention] = useState("");
  const [projectTitle, setProjectTitle] = useState("");
  const [subject, setSubject] = useState("");
  const [refNumber, setRefNumber] = useState("");
  const [date, setDate] = useState(defaultDate);
  const [incomingLetter, setIncomingLetter] = useState("");
  const [bullets, setBullets] = useState("");
  const [tone, setTone] = useState("auto");
  const [selectedRefs, setSelectedRefs] = useState([]);
  const [showRefPicker, setShowRefPicker] = useState(false);

  const [draft, setDraft] = useState("");
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState(null);
  const [copied, setCopied] = useState(false);
  const [savingToLibrary, setSavingToLibrary] = useState(false);
  const [savedNotice, setSavedNotice] = useState(false);
  const [editing, setEditing] = useState(false);

  function toggleRef(id) {
    setSelectedRefs((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  }

  function makeAdder(key) {
    return (value) => addDirectoryValue(key, value);
  }

  async function handleGenerate() {
    if (!bullets.trim()) {
      setError("Add at least a few bullet points for what AQ wants to say.");
      return;
    }
    setGenerating(true);
    setError(null);
    setDraft("");
    try {
      const examples = library
        .filter((e) => selectedRefs.includes(e.id))
        .slice(0, 6)
        .map((e, i) => `EXAMPLE LETTER ${i + 1} (${e.client || "unknown client"} — ${e.project || ""}):\n${e.content.slice(0, 2500)}`)
        .join("\n\n---\n\n");

      const toneNote =
        tone === "cooperative"
          ? "Calibrate tone: cooperative / early-stage, warm but formal."
          : tone === "firm"
          ? "Calibrate tone: firm / escalation register — direct, unhedged, reserving rights, while still closing on a note open to fair resolution."
          : "Calibrate tone automatically based on the context provided (routine, cooperative, or firm) the way AQ naturally would.";

      const userPrompt = `
Draft a business letter with the following details. Output ONLY the letter content starting from the reference/date line through the signature block (do not include the company letterhead graphic — that is added separately). Use plain text with clear line breaks between blocks (recipient address, project/subject lines, salutation, body paragraphs, closing, signature, cc/encl if relevant).

REFERENCE NO.: ${refNumber || "[reference number]"}
DATE: ${date}

RECIPIENT:
${addressee ? addressee + "\n" : ""}${recipientCompany ? recipientCompany + "\n" : ""}${recipientAddress || ""}
${attention ? `KIND ATTN: ${attention}` : ""}

PROJECT: ${projectTitle || "(not specified)"}
SUBJECT: ${subject || "(infer a concise subject line from the bullet points)"}

${incomingLetter.trim() ? `INCOMING LETTER BEING RESPONDED TO:\n${incomingLetter.trim()}\n` : "(This is not a reply to a specific incoming letter — draft it as an outgoing letter.)"}

BULLET POINTS OF WHAT AQ WANTS TO SAY:
${bullets.trim()}

${toneNote}
`.trim();

      const systemPrompt = `${AQ_STYLE_PROFILE}${
        examples ? `\n\nHere are relevant past letters from AQ to calibrate against (match phrasing patterns, structure, and register, but do not copy specific facts/figures from these examples into the new letter):\n\n${examples}` : ""
      }`;

      const settings = loadSettings();
      if (!settings.apiKey) {
        setError("Add your Anthropic API key in Settings (top right) before generating.");
        setGenerating(false);
        return;
      }

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": settings.apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: settings.model || "claude-sonnet-5",
          max_tokens: 1500,
          system: systemPrompt,
          messages: [{ role: "user", content: userPrompt }],
        }),
      });

      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody?.error?.message || `API error (${response.status})`);
      }

      const data = await response.json();
      const text = (data.content || [])
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n")
        .trim();

      if (!text) throw new Error("Empty response");
      setDraft(text);
    } catch (e) {
      setError(e.message || "Something went wrong generating the draft. Please try again.");
    } finally {
      setGenerating(false);
    }
  }

  async function handleCopy() {
    try {
      await navigator.clipboard.writeText(draft);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch (e) {
      // ignore
    }
  }

  async function handleSaveDraftToLibrary() {
    setSavingToLibrary(true);
    try {
      const entry = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        title: refNumber || subject || "Drafted letter",
        client: recipientCompany,
        project: projectTitle,
        date,
        tags: "generated",
        content: draft,
        addedAt: Date.now(),
      };
      await saveLibraryEntry(entry);
      await onSavedToLibrary();
      setSavedNotice(true);
      setTimeout(() => setSavedNotice(false), 2000);
    } catch (e) {
      // ignore
    } finally {
      setSavingToLibrary(false);
    }
  }

  return (
    <div className="grid grid-cols-1 gap-6 lg:grid-cols-5">
      {/* Form panel */}
      <div className="lg:col-span-2 space-y-4">
        <div className="rounded-lg border bg-white p-4 shadow-sm" style={{ borderColor: line, overflow: "visible" }}>
          <SectionLabel>Letterhead</SectionLabel>
          <div className="mt-1.5 flex gap-2">
            <button
              onClick={() => setLetterhead("atl")}
              className="flex-1 rounded-md border px-3 py-1.5 text-xs font-medium"
              style={{
                borderColor: letterhead === "atl" ? ink : line,
                backgroundColor: letterhead === "atl" ? "#EEF1F6" : "white",
                color: letterhead === "atl" ? ink : slate,
              }}
            >
              Al-Tariq Constructors
            </button>
            <button
              onClick={() => setLetterhead("consortium")}
              className="flex-1 rounded-md border px-3 py-1.5 text-xs font-medium"
              style={{
                borderColor: letterhead === "consortium" ? ink : line,
                backgroundColor: letterhead === "consortium" ? "#EEF1F6" : "white",
                color: letterhead === "consortium" ? ink : slate,
              }}
            >
              ATL-GPL Consortium
            </button>
          </div>

          <div className="mt-4 grid grid-cols-2 gap-3">
            <Combobox
              label="Reference no."
              value={refNumber}
              onChange={setRefNumber}
              options={directories.refPrefix}
              placeholder="ATL/236.1/26-054"
              onAddOption={(v) => {
                setRefNumber(v);
                makeAdder("refPrefix")(v);
              }}
            />
            <Field label="Date" value={date} onChange={setDate} placeholder="07.08.2026" />
          </div>
        </div>

        <div className="rounded-lg border bg-white p-4 shadow-sm" style={{ borderColor: line, overflow: "visible" }}>
          <SectionLabel>Recipient</SectionLabel>
          <div className="mt-2 space-y-3">
            <Combobox
              label="Addressee (title / designation)"
              value={addressee}
              onChange={setAddressee}
              options={directories.addressee}
              placeholder="Chief Executive Officer (CEO)"
              onAddOption={(v) => {
                setAddressee(v);
                makeAdder("addressee")(v);
              }}
            />
            <Combobox
              label="Company"
              value={recipientCompany}
              onChange={setRecipientCompany}
              options={directories.recipientCompany}
              placeholder="e.g. Jamshoro Power Company Limited"
              onAddOption={(v) => {
                setRecipientCompany(v);
                makeAdder("recipientCompany")(v);
              }}
            />
            <Combobox
              label="Address"
              value={recipientAddress}
              onChange={setRecipientAddress}
              options={directories.recipientAddress}
              placeholder="Mohra Jabbal, Dadu Road, Jamshoro"
              onAddOption={(v) => {
                setRecipientAddress(v);
                makeAdder("recipientAddress")(v);
              }}
            />
            <Combobox
              label="Kind attention"
              value={attention}
              onChange={setAttention}
              options={directories.kindAttn}
              placeholder="(optional named contact)"
              onAddOption={(v) => {
                setAttention(v);
                makeAdder("kindAttn")(v);
              }}
            />
          </div>
        </div>

        <div className="rounded-lg border bg-white p-4 shadow-sm" style={{ borderColor: line, overflow: "visible" }}>
          <SectionLabel>Subject</SectionLabel>
          <div className="mt-2 space-y-3">
            <Combobox
              label="Project"
              value={projectTitle}
              onChange={setProjectTitle}
              options={directories.projectTitle}
              placeholder="(optional)"
              onAddOption={(v) => {
                setProjectTitle(v);
                makeAdder("projectTitle")(v);
              }}
            />
            <Combobox
              label="Subject line"
              value={subject}
              onChange={setSubject}
              options={directories.subject}
              placeholder="(optional — will be inferred if left blank)"
              onAddOption={(v) => {
                setSubject(v);
                makeAdder("subject")(v);
              }}
            />
          </div>
        </div>

        <div className="rounded-lg border bg-white p-4 shadow-sm" style={{ borderColor: line }}>
          <SectionLabel>Content</SectionLabel>
          <div className="mt-2 space-y-3">
            <div>
              <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>Incoming letter (optional)</label>
              <textarea
                value={incomingLetter}
                onChange={(e) => setIncomingLetter(e.target.value)}
                rows={4}
                placeholder="Paste the letter being replied to, if any..."
                className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
                style={{ borderColor: line, fontFamily: "Tinos" }}
              />
            </div>
            <div>
              <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>Bullet points — what AQ wants to say</label>
              <textarea
                value={bullets}
                onChange={(e) => setBullets(e.target.value)}
                rows={6}
                placeholder={"- Reference our letter dated ...\n- Amount of Rs. ... is still overdue\n- Request release within 7 days"}
                className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
                style={{ borderColor: line }}
              />
            </div>
          </div>
        </div>

        <div className="rounded-lg border bg-white p-4 shadow-sm" style={{ borderColor: line }}>
          <SectionLabel>Tone</SectionLabel>
          <div className="mt-1.5 flex gap-2">
            {[
              { id: "auto", label: "Auto" },
              { id: "cooperative", label: "Cooperative" },
              { id: "firm", label: "Firm / escalation" },
            ].map((t) => (
              <button
                key={t.id}
                onClick={() => setTone(t.id)}
                className="flex-1 rounded-md border px-2.5 py-1.5 text-xs font-medium"
                style={{
                  borderColor: tone === t.id ? brass : line,
                  backgroundColor: tone === t.id ? "#FBF3E4" : "white",
                  color: tone === t.id ? brass : slate,
                }}
              >
                {t.label}
              </button>
            ))}
          </div>

          <div className="mt-4">
            <button
              onClick={() => setShowRefPicker((v) => !v)}
              className="flex w-full items-center justify-between rounded-md border px-3 py-2 text-xs font-medium"
              style={{ borderColor: line, color: slate }}
            >
              <span>
                Reference {selectedRefs.length > 0 ? `(${selectedRefs.length} selected)` : "from style library"}
              </span>
              <ChevronDown size={14} style={{ transform: showRefPicker ? "rotate(180deg)" : "none" }} />
            </button>
            {showRefPicker && (
              <div className="mt-2 max-h-48 space-y-1 overflow-y-auto rounded-md border p-2" style={{ borderColor: line }}>
                {library.length === 0 ? (
                  <div style={{ color: slateLight, fontSize: "11.5px" }}>No letters in your library yet.</div>
                ) : (
                  library.map((e) => (
                    <label key={e.id} className="flex cursor-pointer items-start gap-2 rounded px-1.5 py-1 text-xs hover:bg-gray-50">
                      <input
                        type="checkbox"
                        checked={selectedRefs.includes(e.id)}
                        onChange={() => toggleRef(e.id)}
                        className="mt-0.5"
                      />
                      <span style={{ color: charcoal }}>
                        {e.title} {e.client ? <span style={{ color: slateLight }}>— {e.client}</span> : null}
                      </span>
                    </label>
                  ))
                )}
              </div>
            )}
          </div>
        </div>

        {error && (
          <div className="flex items-center gap-1.5 rounded-md border px-3 py-2" style={{ borderColor: maroon, color: maroon, fontSize: "12px" }}>
            <AlertCircle size={13} /> {error}
          </div>
        )}

        <button
          onClick={handleGenerate}
          disabled={generating}
          className="flex w-full items-center justify-center gap-2 rounded-md py-2.5 text-sm font-semibold text-white shadow-sm"
          style={{ backgroundColor: ink, opacity: generating ? 0.7 : 1 }}
        >
          {generating ? <Loader2 size={16} className="animate-spin" /> : <Send size={15} />}
          {generating ? "Drafting..." : "Generate draft"}
        </button>
      </div>

      {/* Preview panel */}
      <div className="lg:col-span-3">
        <div className="mb-2 flex items-center justify-between">
          <SectionLabel>Preview</SectionLabel>
          {draft && (
            <div className="flex gap-2">
              <button
                onClick={() => setEditing((v) => !v)}
                className="rounded-md border px-2.5 py-1 text-xs font-medium"
                style={{ borderColor: line, color: slate }}
              >
                {editing ? "Done editing" : "Edit"}
              </button>
              <button
                onClick={handleSaveDraftToLibrary}
                disabled={savingToLibrary}
                className="flex items-center gap-1 rounded-md border px-2.5 py-1 text-xs font-medium"
                style={{ borderColor: line, color: slate }}
              >
                {savedNotice ? <Check size={12} /> : <BookOpen size={12} />}
                {savedNotice ? "Saved" : "Save to library"}
              </button>
              <button
                onClick={handleCopy}
                className="flex items-center gap-1 rounded-md px-2.5 py-1 text-xs font-medium text-white"
                style={{ backgroundColor: copied ? "#3E7D52" : brass }}
              >
                {copied ? <Check size={12} /> : <Copy size={12} />}
                {copied ? "Copied" : "Copy"}
              </button>
            </div>
          )}
        </div>

        <div
          className="rounded-md p-8 shadow-lg"
          style={{ backgroundColor: parchment, border: `1px solid ${line}`, minHeight: "600px" }}
        >
          {letterhead === "atl" ? <LetterheadATL /> : <LetterheadConsortium />}

          <div className="mt-6" style={{ fontFamily: "Tinos", color: charcoal, fontSize: "14px", lineHeight: 1.75 }}>
            {generating ? (
              <div className="flex items-center gap-2 py-16 justify-center" style={{ color: slate, fontFamily: "Inter" }}>
                <Loader2 size={18} className="animate-spin" /> Drafting in AQ's voice...
              </div>
            ) : draft ? (
              editing ? (
                <textarea
                  value={draft}
                  onChange={(e) => setDraft(e.target.value)}
                  rows={22}
                  className="w-full resize-y rounded border-none bg-transparent p-2 outline-none"
                  style={{ fontFamily: "Tinos", fontSize: "14px", lineHeight: 1.75, backgroundColor: parchmentDark }}
                />
              ) : (
                <div style={{ whiteSpace: "pre-wrap" }}>{draft}</div>
              )
            ) : (
              <div className="flex flex-col items-center gap-2 py-20 text-center" style={{ color: slateLight, fontFamily: "Inter" }}>
                <FileText size={26} />
                <div style={{ fontSize: "13px" }}>
                  Fill in the details and generate a draft — it will appear here on AQ's letterhead.
                </div>
              </div>
            )}
          </div>

          {letterhead === "consortium" && <LetterFooterConsortium />}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Settings modal — API key + model, stored in localStorage only (never sent
// anywhere but directly from your browser to api.anthropic.com)
// ---------------------------------------------------------------------------
function SettingsModal({ onClose }) {
  const initial = loadSettings();
  const [apiKey, setApiKey] = useState(initial.apiKey || "");
  const [model, setModel] = useState(initial.model || "claude-sonnet-5");
  const [saved, setSaved] = useState(false);

  function handleSave() {
    saveSettings({ apiKey: apiKey.trim(), model: model.trim() || "claude-sonnet-5" });
    setSaved(true);
    setTimeout(() => {
      setSaved(false);
      onClose();
    }, 700);
  }

  return (
    <div className="fixed inset-0 z-30 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div
        className="w-full max-w-md rounded-lg bg-white p-5 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between">
          <h3 style={{ color: ink, fontWeight: 700, fontSize: "15px" }}>Settings</h3>
          <button onClick={onClose} style={{ color: slateLight }}>
            <X size={18} />
          </button>
        </div>

        <div className="space-y-3">
          <div>
            <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>Anthropic API key</label>
            <input
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder="sk-ant-..."
              className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
              style={{ borderColor: line, fontFamily: "JetBrains Mono" }}
            />
            <p style={{ color: slateLight, fontSize: "11px", marginTop: "4px" }}>
              Stored only in this browser's local storage. Sent directly to api.anthropic.com — never to any
              other server. Get a key at{" "}
              <a href="https://console.anthropic.com" target="_blank" rel="noreferrer" style={{ color: brass }}>
                console.anthropic.com
              </a>
              .
            </p>
          </div>

          <div>
            <label style={{ color: slate, fontSize: "11.5px", fontWeight: 500 }}>Model</label>
            <input
              value={model}
              onChange={(e) => setModel(e.target.value)}
              placeholder="claude-sonnet-5"
              className="mt-1 w-full rounded-md border px-3 py-2 text-sm outline-none"
              style={{ borderColor: line, fontFamily: "JetBrains Mono" }}
            />
            <p style={{ color: slateLight, fontSize: "11px", marginTop: "4px" }}>
              Defaults to claude-sonnet-5. Change only if you know the exact model ID you want.
            </p>
          </div>
        </div>

        <div className="mt-5 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-md px-3.5 py-2 text-sm" style={{ color: slate }}>
            Cancel
          </button>
          <button
            onClick={handleSave}
            className="flex items-center gap-1.5 rounded-md px-4 py-2 text-sm font-medium text-white"
            style={{ backgroundColor: ink }}
          >
            {saved ? <Check size={14} /> : null}
            {saved ? "Saved" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

function SectionLabel({ children }) {
  return (
    <div style={{ color: slate, fontSize: "11px", fontWeight: 700, letterSpacing: "0.6px", textTransform: "uppercase" }}>
      {children}
    </div>
  );
}
