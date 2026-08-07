const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, HeadingLevel,
  AlignmentType, ShadingType, BorderStyle, LevelFormat, convertInchesToTwip,
  Table, TableRow, TableCell, WidthType, TableLayoutType,
} = require("docx");

// ---------------------------------------------------------------- palette ---
const SC_BLUE = "0061C7";  // primary
const NAVY = "2C3A87";     // category bars, subcategory heading text
const TINT = "D9E7F7";     // subcategory box background
const GREEN = "92E773";    // banner eyebrow
const BLACK = "000000";

const NUM = "bulletList";

// ------------------------------------------------------------------ links ---
const L = {
  onlineBanking: "https://www.sc.com/sg/bank-with-us/online-banking/",
  scMobile: "https://www.sc.com/sg/bank-with-us/mobile-banking-services/standard-chartered-mobile/",
  appStore: "https://itunes.apple.com/sg/app/sc-mobile-singapore-breeze/id367337298?mt=8",
  googlePlay: "https://play.google.com/store/apps/details?id=air.app.scb.breeze.android.main.sg.prod&hl=en",
  activateOnline: "https://forms.online.standardchartered.com/public_website/singapore/breezeform/index.html",
  overseasOnline: "https://online.forms.standardchartered.com/public_website/singapore/breezeform/index.html",
  activationFaqs: "https://www.sc.com/sg/help/faqs/activate-new-card",
  overseasFaqs: "https://www.sc.com/sg/help/faqs/activate-overseas-card-usage",
  absActivation: "http://www.abs.org.sg/financial_faq_payment_p2c.php",
  absOverseas: "http://www.abs.org.sg/financial_faq_payment_p2a.php",
};

// Rich-text mini-DSL: a string is plain text, ["label", url] is a hyperlink.
function runs(parts, { bold = false, size = 22, color = BLACK } = {}) {
  return parts.map((p) =>
    Array.isArray(p)
      ? new ExternalHyperlink({
          link: p[1],
          children: [new TextRun({ text: p[0], size, color: SC_BLUE, underline: {} })],
        })
      : new TextRun({ text: p, bold, size, color })
  );
}

// ---------------------------------------------------------------- banner ----
// A4 (11906 twips) less the 1080-twip side margins set on the section.
const CONTENT_W = 11906 - 1080 * 2; // 9746
const BANNER_PAD = 260;             // single padding value, all four sides
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "auto" };

// Both banner lines live in ONE table cell whose margins supply the only
// horizontal padding, so the two lines cannot drift apart: neither paragraph
// carries an indent of its own.
function banner(title) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    layout: TableLayoutType.FIXED,
    margins: { top: BANNER_PAD, bottom: BANNER_PAD, left: BANNER_PAD, right: BANNER_PAD },
    borders: {
      top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER,
      insideHorizontal: NO_BORDER, insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: SC_BLUE, color: "auto" },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0, line: 260 },
                indent: { left: 0, right: 0 },
                children: [
                  new TextRun({
                    text: "Standard Chartered Singapore",
                    bold: true, smallCaps: true, size: 20, color: GREEN, characterSpacing: 80,
                  }),
                ],
              }),
              // Heading 1 lives inside the banner so the title is TOC-navigable
              // without being repeated as a separate visible heading below it.
              new Paragraph({
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 60, after: 0, line: 420 },
                indent: { left: 0, right: 0 },
                children: [new TextRun({ text: title, bold: true, size: 40, color: "FFFFFF" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// -------------------------------------------------------------- structure ---
function categoryBar(name) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    shading: { type: ShadingType.CLEAR, fill: NAVY, color: "auto" },
    spacing: { before: 360, after: 180, line: 320 },
    indent: { left: 160, right: 160 },
    children: [new TextRun({ text: name, bold: true, size: 28, color: "FFFFFF" })],
  });
}

const edge = { style: BorderStyle.SINGLE, size: 8, space: 6, color: SC_BLUE };

function subcategoryBox(category, name, url) {
  const shade = { type: ShadingType.CLEAR, fill: TINT, color: "auto" };
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_3,
      shading: shade,
      border: { top: edge, left: edge, right: edge },
      spacing: { before: 240, after: 0, line: 300 },
      indent: { left: 160, right: 160 },
      children: [new TextRun({ text: `${category} › ${name}`, bold: true, size: 24, color: NAVY })],
    }),
    new Paragraph({
      shading: shade,
      border: { bottom: edge, left: edge, right: edge },
      spacing: { before: 0, after: 200, line: 260 },
      indent: { left: 160, right: 160 },
      children: [
        new TextRun({ text: "Source: ", italics: true, size: 18, color: NAVY }),
        new ExternalHyperlink({
          link: url,
          children: [new TextRun({ text: url, italics: true, size: 18, color: SC_BLUE, underline: {} })],
        }),
      ],
    }),
  ];
}

function qa(n, q, aParts) {
  const out = [
    new Paragraph({
      spacing: { before: 220, after: 60 },
      children: [
        new TextRun({ text: `Q${n}: `, bold: true, size: 22, color: SC_BLUE }),
        ...runs(Array.isArray(q) ? q : [q], { bold: true }),
      ],
    }),
  ];
  aParts.forEach((part, i) => {
    if (part && part.bullets) {
      part.bullets.forEach((b) => {
        out.push(new Paragraph({
          numbering: { reference: NUM, level: 0 },
          spacing: { before: 20, after: 20 },
          children: runs(b), // b is always an array of DSL parts
        }));
      });
    } else {
      out.push(new Paragraph({
        spacing: { before: i === 0 ? 0 : 80, after: 60 },
        children: [
          ...(i === 0 ? [new TextRun({ text: `A${n}: `, bold: true, size: 22, color: SC_BLUE })] : []),
          ...runs(Array.isArray(part) ? part : [part]),
        ],
      }));
    }
  });
  return out;
}
const B = (bullets) => ({ bullets });

// ---------------------------------------------------------------- content ---
const CATEGORIES = [
  {
    name: "Card Services",
    subs: [
      {
        name: "Activate Card",
        url: "https://www.sc.com/sg/help/activate-card/",
        qas: [
          ["What is the status of a new Standard Chartered ATM, debit or credit card when it arrives?",
           [["Every ATM, debit and credit card Standard Chartered Singapore sends out arrives blocked for added security. The card must be activated online or via SMS before it can be used."]]],

          ["What is required to activate a Standard Chartered ATM, debit or credit card?",
           ["Activation of a Standard Chartered ATM, debit or credit card requires the following:",
            B([["A mobile phone number registered with the Bank. This number is also required to enable and disable overseas ATM cash withdrawal."],
               ["A One-Time Password (OTP), which is sent to the registered mobile phone and used to complete card activation."]])]],

          ["How is a registered mobile number updated before activating a Standard Chartered card?",
           [["To update a registered mobile number, log in to ", ["Online Banking", L.onlineBanking], " or ", ["SC Mobile", L.scMobile], " and go to “Update Profile Details”."]]],

          ["Who activates a Standard Chartered supplementary card?",
           [["Principal cardholders activate supplementary cards. A supplementary cardholder does not perform the activation."]]],

          ["How is a Standard Chartered card activated through Online Banking or SC Mobile?",
           [["A Standard Chartered ATM, debit or credit card can be activated and its PIN set through ", ["Online Banking", L.onlineBanking], " or ", ["SC Mobile", L.scMobile], " using these steps:"],
            B([["Step 1: Log in to Online Banking or SC Mobile."],
               ["Step 2: Go to “Help & Services”. In SC Mobile, tap the Profile icon at the top left of the screen to reach “Help & Services”."],
               ["Step 3: Select “Credit Card Activation & PIN Set” or “Debit/ATM Card Activation & PIN Set”."],
               ["Step 4: Follow the on-screen instructions."],
               ["Step 5: On completion, the card is activated immediately and the PIN is set."]])]],

          ["Where can a Standard Chartered card be activated online without logging in to Online Banking or SC Mobile?",
           [["A Standard Chartered card can also be activated through the Bank’s online card activation form at ", ["forms.online.standardchartered.com", L.activateOnline], ". Further answers on card activation are published at ", ["sc.com/sg/help/faqs/activate-new-card", L.activationFaqs], ", and general guidance is available on the Association of Banks in Singapore website at ", ["abs.org.sg", L.absActivation], "."]]],

          ["Where can SC Mobile be downloaded?",
           [["SC Mobile is Standard Chartered’s mobile banking app and is available free of charge on the ", ["App Store (iOS)", L.appStore], " and ", ["Google Play (Android)", L.googlePlay], ". App Store is a trademark of Apple Inc. registered in the US and other countries. Android is a trademark of Google Inc."]]],

          ["How is a Standard Chartered ATM or debit card activated by SMS?",
           [["To activate a Standard Chartered ATM or debit card by SMS, send the following to 75722: ACT<space>last four digits of the card number. For example: ACT 5432."]]],

          ["How is a Standard Chartered credit card activated by SMS?",
           [["To activate a Standard Chartered credit card by SMS, send the following to 75722: ACT<space>name on card<space>last four digits of the card number. For example: ACT M H Tan 5432."]]],

          ["Which overseas card functions are disabled by default on Standard Chartered cards?",
           ["Standard Chartered disables the overseas card usage function by default on all new, existing, replacement and renewed ATM, credit and debit cards. The functions affected are:",
            B([["Overseas ATM cash withdrawal."],
               ["Overseas credit/debit card magnetic stripe transactions."]])]],

          ["Are there exceptions to the default disabling of overseas card usage on Standard Chartered cards?",
           ["The overseas card usage function is not disabled by default in these cases:",
            B([["The overseas ATM cash withdrawal function had previously been enabled on the card. This also applies to overseas card usage on renewed or replacement cards carrying the same card number."],
               ["The cardholder was notified that their existing card or cards will remain enabled for overseas usage."]])]],

          ["Can a Standard Chartered card still be used overseas while the overseas card usage function is disabled?",
           [["Yes. EMV chip transactions are excluded from the overseas card usage function, so the card can still be used overseas at EMV chip-enabled POS terminals and ATMs even when the overseas card usage function is disabled. “POS” stands for Point-of-Sale and “EMV” stands for Europay, Mastercard and Visa."]]],

          ["How is overseas card usage enabled or disabled online for a Standard Chartered card?",
           [["Overseas card usage can be enabled or disabled using the Bank’s online form at ", ["online.forms.standardchartered.com", L.overseasOnline], ", or through ", ["SC Mobile", L.scMobile], ". This can only be done after the card has been activated. Further answers on overseas card usage are published at ", ["sc.com/sg/help/faqs/activate-overseas-card-usage", L.overseasFaqs], ", and general guidance is available on the Association of Banks in Singapore website at ", ["abs.org.sg", L.absOverseas], "."]]],

          ["How is overseas card usage enabled by SMS for a Standard Chartered ATM or debit card?",
           [["To enable overseas card usage on a Standard Chartered ATM or debit card, send the following to 75722: AOT<space>last four digits of the card number. For example: AOT 5432. To disable it, send the same message with the keyword “DOT” in place of “AOT”."]]],

          ["How is overseas card usage enabled by SMS for a Standard Chartered credit card?",
           [["To enable overseas card usage on a Standard Chartered credit card, send the following to 75722: AOT<space>name on card<space>last four digits of the card number. For example: AOT M H Tan 5432. To disable it, send the same message with the keyword “DOT” in place of “AOT”."]]],

          ["How long does overseas card usage stay enabled when it is activated by SMS on a Standard Chartered card?",
           [["Overseas card usage enabled by SMS is valid for 12 months from the date of activation. To enable overseas usage perpetually instead, replace the keyword “AOT” with “AOTP” in the SMS. An expiry date for overseas usage can be set if the function is enabled through the Bank’s online form at ", ["online.forms.standardchartered.com", L.overseasOnline], ". Confirmation is sent by SMS once the card or cards have been successfully enabled."]]],

          ["What is the risk of enabling overseas usage on a Standard Chartered debit or credit card?",
           [["The risk of unauthorised transactions on debit and credit cards is higher when the cards are activated for overseas use, because the data encoded on a card’s magnetic stripe may be easily read and replicated. Standard Chartered recommends signing up for SMS alerts on card transactions to guard against this risk, and reporting any unauthorised transaction immediately."]]],

          ["Does Standard Chartered still issue a physical PIN mailer for new cards?",
           [["No. In line with the Bank’s commitment towards a more sustainable future, Standard Chartered no longer issues a physical PIN for any ATM, debit card or credit card issuance. The PIN is set by the cardholder through SC Mobile, or at a Standard Chartered ATM for ATM and debit cards."]]],

          ["How is a card PIN set or changed using SC Mobile?",
           [["A Standard Chartered card PIN can be set or changed in ", ["SC Mobile", L.scMobile], " as follows:"],
            B([["Step 1: Log in to SC Mobile."],
               ["Step 2: Go to the Service icon at the bottom right of the screen and navigate to “Digital Services”."],
               ["Step 3: Go to the “Card Management” section and select the applicable option. If the card has not been activated, select “Debit/ATM Card Activation & PIN set” or “Credit Card Activation and PIN set”. If the card has already been activated, select “Debit/ATM Card PIN change” or “Credit Card PIN change”."],
               ["Step 4: Follow the on-screen instructions to set the PIN."]])]],

          ["How is a PIN set at a Standard Chartered ATM?",
           ["Setting a PIN at a Standard Chartered ATM is available for ATM and debit cards only. The steps are:",
            B([["Step 1: Insert the ATM or debit card into a Standard Chartered ATM."],
               ["Step 2: Select “Set/Reset PIN & Card Activation for ATM/Debit Card”."],
               ["Step 3: Select “Request OTP” to have a one-time password sent to the registered mobile number."],
               ["Step 4: Enter the OTP."],
               ["Step 5: Set a new 5-digit PIN."]])]],
        ],
      },
      {
        name: "Lost Card",
        url: "https://www.sc.com/sg/help/lost-card/",
        qas: [
          ["When should a Standard Chartered card be reported as lost or stolen?",
           [["A Standard Chartered card should be reported if it is lost or stolen, or if its PIN has been disclosed to a third person. The report can be made by logging in to ", ["Online Banking", L.onlineBanking], " or ", ["SC Mobile", L.scMobile], "."]]],

          ["How is a lost or stolen Standard Chartered card reported through Online Banking or SC Mobile?",
           [["A lost or stolen Standard Chartered card is reported through ", ["Online Banking", L.onlineBanking], " or ", ["SC Mobile", L.scMobile], " as follows:"],
            B([["Step 1: Log in to Online Banking or SC Mobile."],
               ["Step 2: Go to “Help & Services”."],
               ["Step 3: Select “Report Lost/Stolen Card”."],
               ["Step 4: Follow the on-screen instructions."],
               ["Step 5: On completion, the card is blocked and a replacement card is sent to the registered mailing address."]])]],

          ["Is a replacement card issued automatically after a Standard Chartered card is reported lost or stolen?",
           [["Yes. Once a lost or stolen card report is completed through Online Banking or SC Mobile, the card is blocked and a replacement card is sent to the cardholder’s registered mailing address."]]],

          ["How is a Standard Chartered credit or debit card PIN reset?",
           [["A Standard Chartered credit or debit/ATM card PIN is reset through ", ["Online Banking", L.onlineBanking], " or ", ["SC Mobile", L.scMobile], " as follows:"],
            B([["Step 1: Log in to Online Banking or SC Mobile."],
               ["Step 2: Go to “Help & Services”."],
               ["Step 3: Select “Credit Card PIN Change” or “Debit/ATM Card PIN Change”."],
               ["Step 4: Follow the on-screen instructions."],
               ["Step 5: On completion, the credit or debit/ATM card PIN is changed."]])]],

          ["Is a cardholder liable for charges on a Standard Chartered card that is lost or stolen?",
           [["The cardholder, or the company if the card is a corporate card, is liable for all unauthorised card transactions effected before Standard Chartered is notified of the card’s loss, theft or PIN disclosure. Full details of the cardholder’s responsibilities and the Bank’s responsibilities are set out in the Standard Chartered Cardmembers Agreement."]]],
        ],
      },
    ],
  },
  {
    name: "Auto Financing",
    subs: [
      {
        name: "Interest Computation and Fees & Charges",
        url: "https://www.sc.com/sg/help/auto-financing/",
        qas: [
          ["How is interest charged on a Standard Chartered auto financing facility?",
           [["Market practice is to charge what is known as a “flat” or “applied” interest rate. This rate is derived by assuming that the Amount Financed is constant throughout the hire period. However, the Amount Financed reduces over time, so the Effective Interest Rate (EIR) is higher than the flat or applied interest rate."]]],

          ["What does a worked example of flat interest rate calculation on auto financing look like?",
           ["Standard Chartered provides the following example of a flat interest rate calculation for auto financing:",
            B([["Amount financed: S$50,000."],
               ["Hire period: 60 months (5 years)."],
               ["Flat interest rate: 2.78%."],
               ["EIR: 5.24%."],
               ["Total interest payable: 2.78% x $50,000 x 5 = $6,950."],
               ["Rental charges / monthly instalment = $(50,000 + 6,950) / 60 = $949.17."]])]],

          ["How does the split between interest and principal change over an auto financing instalment schedule?",
           [["The amount of interest paid through each monthly instalment reduces over time, while the amount of principal paid through each monthly instalment increases, so that the total monthly instalment amount stays unchanged. This is in line with market practice and is commonly known as the ‘Rule of 78’."]]],

          ["What early completion fee applies if an auto financing vehicle purchase is completed before the hire period expires?",
           ["If the purchase of the vehicle is completed before the expiry of the hire period, an early completion fee is calculated as follows:",
            B([["If the date of completion is less than 12 months from the commencement date, the early completion fee is 20% of outstanding interest + 1.5% of balance payable."],
               ["If the date of completion is 12 months or more from the commencement date, the early completion fee is 20% of outstanding interest + 1.5% of balance payable."]])]],

          ["How is “Outstanding Interest” calculated for a Standard Chartered auto financing early completion fee?",
           ["“Outstanding Interest” at any time is calculated as [n(n+1) x TI] / [N(N+1)], where:",
            B([["n is the number of months remaining in the hire period."],
               ["N is the total number of months in the hire period."],
               ["TI is the Total Interest."]])]],
        ],
      },
      {
        name: "Other Useful Information",
        url: "https://www.sc.com/sg/help/auto-financing/",
        qas: [
          ["Which channels can be used to pay a Standard Chartered Auto Financing bill?",
           ["Auto Financing bills can be paid through any of the following channels:",
            B([["GIRO. To apply, mail the completed GIRO form to Standard Chartered."],
               ["Online Banking, or internet transfer from any other bank."],
               ["NETS at AXS stations located island wide."],
               ["Cheque. Write the Hire Purchase Agreement Number, NRIC and contact number on the reverse side of the cheque. Cheques should be crossed and made payable to the repayment account, and may be deposited at any Standard Chartered branch."],
               ["Cash payment at a Standard Chartered Cash Deposit Machine."]])]],

          ["How is early redemption of a Standard Chartered auto financing facility performed?",
           [["An early redemption request can be submitted through a dealer if a new vehicle is being purchased. A redemption quote can otherwise be requested from Standard Chartered directly."]]],

          ["Can the monthly due date on a Standard Chartered auto financing facility be changed?",
           [["Yes. A change of due date is allowed, but it is subject to approval and to the total tenure of the facility, which may be bound by regulatory guidelines."]]],
        ],
      },
    ],
  },
];

// ------------------------------------------------------------------ build ---
const children = [
  banner("Standard Chartered Singapore Help Centre FAQ"),
  new Paragraph({ spacing: { before: 0, after: 0, line: 240 }, children: [] }), // spacer below banner
];
let counter = 0; // continuous across the whole document — never resets
for (const cat of CATEGORIES) {
  children.push(categoryBar(cat.name));
  for (const sub of cat.subs) {
    children.push(...subcategoryBox(cat.name, sub.name, sub.url));
    for (const [q, a] of sub.qas) children.push(...qa(++counter, q, a));
  }
}
console.log("total Q&A:", counter);

const doc = new Document({
  title: "SCB Bank_Help FAQ",
  description: "Standard Chartered Singapore Help Centre FAQ — RAG-ready",
  creator: "Standard Chartered Singapore FAQ extract",
  numbering: {
    config: [{
      reference: NUM,
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.2) } } },
      }],
    }],
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: BLACK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Calibri", size: 40, bold: true, color: "FFFFFF" } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Calibri", size: 28, bold: true, color: "FFFFFF" } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Calibri", size: 24, bold: true, color: NAVY } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4 — CONTENT_W is derived from this
        margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Help_FAQ.docx", buf);
  console.log("written", buf.length);
});
