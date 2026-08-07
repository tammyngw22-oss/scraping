const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

const U = {
  overview: "https://www.sc.com/sg/credit-cards/",
  simplyCash: "https://www.sc.com/sg/credit-cards/simply-cash-credit-card/",
  beyond: "https://www.sc.com/sg/credit-cards/beyond-credit-card/",
  journey: "https://www.sc.com/sg/credit-cards/journey-credit-card/",
  smart: "https://www.sc.com/sg/credit-cards/smart-credit-card/",
  visaInfinite: "https://www.sc.com/sg/credit-cards/visa-infinite-card/",
  priorityVI: "https://www.sc.com/sg/credit-cards/priority-visa-infinite-card/",
  rewardsPlus: "https://www.sc.com/sg/credit-cards/rewards-plus-credit-card/",
  platinum: "https://www.sc.com/sg/credit-cards/platinum-visa-mastercard-credit-card/",
  pruPlatinum: "https://www.sc.com/sg/credit-cards/prudential-platinum-card/",
  pruSignature: "https://www.sc.com/sg/credit-cards/prudential-visa-signature-card/",
  nus: "https://www.sc.com/sg/credit-cards/nus-alumni-platinum-card/",
};

// Links referenced inside answers.
const L = {
  simplyCashTnc: "https://www.sc.com/sg/terms-and-conditions/simply-cash-credit-card-terms-and-conditions/",
  simplyCashTncPdf: "https://av.sc.com/sg/content/docs/sg-simply-cash-card-terms-and-conditions.pdf",
  journeyTnc: "https://av.sc.com/sg/content/docs/sg-journey-credit-card-terms-and-conditions.pdf",
  rewardsPlusTnc: "https://www.sc.com/sg/rewardsplustncs",
  pbviTnc: "https://av.sc.com/sg/content/docs/sg-pbvi-tnc-full.pdf",
  beyondTnc: "https://www.sc.com/sg/terms-and-conditions/beyond-credit-card-tnc/",
  rewards360: "https://www.sc.com/sg/rewards-programmes/360-rewards/",
  rewardsRedeem: "https://www.sc.com/sg/rewards-programmes/360-rewards-redemption/",
  pruAccess: "https://pruaccess.prudential.com.sg/pruaccess_sg/",
  iPay: "https://ipay.prudential.com.sg/ipay/",
  creditBureau: "https://www.creditbureau.com.sg",
  pricingGuide: "https://www.sc.com/sg/pricing-guide/",
  cards: "https://www.sc.com/sg/credit-cards/",
};

// Boilerplate shared verbatim across the card pages.
const FINANCE = "The Effective Interest Rate (EIR) is 27.9% per annum (minimum). If payment is not made in full by the due date, finance charges are calculated on a daily basis at 0.076% from the respective transaction dates for all transactions to the date the payment is received.";
const CASH_ADV = "The cash advance fee per transaction is 8% on the Cash Advance, subject to a minimum fee of $15, plus Finance Charges at 0.082% per day on the amount withdrawn from the date of the transaction until the date of full payment.";
const LATE = "A late payment charge of S$100 is charged if the minimum payment due is not received by the due date.";

const CATEGORIES = [
  // ------------------------------------------------------------------------
  {
    name: "Choosing a Credit Card",
    subs: [
      {
        name: "Card Range Overview",
        url: U.overview,
        qas: [
          ["Which Standard Chartered Singapore credit cards are positioned for travel and miles?",
           ["Standard Chartered Singapore lists three cards under its travel positioning:",
            B([["SC Visa Infinite — 3 miles per S$1 on overseas spend and 1.4 miles per S$1 on local spend, both with a minimum S$2,000 spend."],
               ["SC Journey — 3 miles per S$1 on rides, food deliveries and groceries, and 2 miles per S$1 on all overseas spend with no cap."],
               ["SC Beyond — up to 8 miles per S$1 on overseas dining, plus complimentary Business Class upgrades."]])]],

          ["Which Standard Chartered Singapore credit cards are positioned for cashback?",
           ["Standard Chartered Singapore lists two cards under its cashback positioning:",
            B([["SC Simply Cash — 1.5% unlimited cashback with no minimum spend."],
               ["SC Smart — up to 10% cashback on streaming, dining and commuting, with no cap on the cashback that can be earned."]])]],

          ["What welcome offer applies to the SC Visa Infinite Credit Card?",
           [["The SC Visa Infinite Credit Card offers up to 50,000 Welcome Miles with a minimum spend of S$2,000 within 60 days of card approval. An annual fee of S$550 (excluding GST) applies. Terms and conditions apply."]]],

          ["What welcome offer applies to the SC Journey Credit Card?",
           [["The SC Journey Credit Card offers up to 30,000 Welcome Miles with a minimum spend of S$800 within 60 days of card approval. An annual fee of S$180 (excluding GST) applies. Terms and conditions apply."]]],

          ["What welcome offer applies to the SC Simply Cash and SC Smart Credit Cards?",
           [["Both the SC Simply Cash Credit Card and the SC Smart Credit Card offer a welcome bonus worth S$710 for new cardholders: a Samsonite MODUS Spinner 25” luggage and S$100 cashback. Terms and conditions apply."]]],

          ["Which Standard Chartered Singapore credit cards include complimentary airport lounge access?",
           ["Lounge access differs by card:",
            B([["SC Visa Infinite — six complimentary Priority Pass visits at over 1,700 airport lounges and travel experiences."],
               ["SC Journey — two complimentary Priority Pass visits at over 1,700 airport lounges and travel experiences."],
               ["SC Beyond — unlimited Priority Pass visits for the cardholder and Supplementary Cardholders, plus six complimentary visits for guests per year."],
               ["SC Smart — Visa SmartDelay gives complimentary access to over 1,600 airport lounges and travel experiences when a flight is delayed."]])]],

          ["What complimentary travel insurance coverage do the Standard Chartered Singapore travel cards provide?",
           ["Complimentary travel insurance coverage differs by card:",
            B([["SC Visa Infinite — coverage of up to S$1,000,000 for two."],
               ["SC Journey — coverage of up to S$500,000."],
               ["SC Beyond — coverage of up to USD 500,000, covering travel inconveniences such as delays, missed connections, loss of baggage, medical expenses and medical evacuation."]])]],

          ["Which Caltex benefit applies across the Standard Chartered Singapore credit card range?",
           [["Cardholders of the SC Visa Infinite, SC Simply Cash, SC Journey and SC Smart Credit Cards receive six S$5 Caltex vouchers, worth S$30 in total, with The Good Life upon card approval. The Good Life also provides dining, entertainment and travel deals at more than 3,000 outlets across Asia."]]],
        ],
      },
    ],
  },

  // ------------------------------------------------------------------------
  {
    name: "Cashback Cards",
    subs: [
      {
        name: "Simply Cash Credit Card",
        url: U.simplyCash,
        qas: [
          ["Are there different cashback tiers on the Standard Chartered Simply Cash Credit Card?",
           [["There are no cashback tiers. The Simply Cash Credit Card offers 1.5% cashback on all eligible spends."]]],

          ["Is a minimum spend required to earn cashback on the Standard Chartered Simply Cash Credit Card?",
           [["There is no minimum spend required to be eligible for cashback on the Simply Cash Credit Card."]]],

          ["When is cashback credited on the Standard Chartered Simply Cash Credit Card?",
           [["The cashback for eligible spends in each statement cycle is credited into the Simply Cash Credit Card in the next statement month."]]],

          ["Is there a cap on the cashback earned on the Standard Chartered Simply Cash Credit Card?",
           [["There is no cap on the cashback amount that can be earned. The Simply Cash Credit Card gives 1.5% cashback on all eligible spends."]]],

          ["Which transactions are excluded from earning cashback on the Standard Chartered Simply Cash Credit Card?",
           [["The list of transactions that are not eligible to earn cashback on the Simply Cash Credit Card is published in the card’s terms and conditions at ", ["sc.com/sg/terms-and-conditions/simply-cash-credit-card-terms-and-conditions", L.simplyCashTnc], ". The full terms are also available as a PDF at ", ["av.sc.com", L.simplyCashTncPdf], "."]]],

          ["What is the welcome gift on the Standard Chartered Simply Cash Credit Card?",
           [["From now till 31 August 2026, applicants receive a Samsonite MODUS Spinner 25” worth S$610 and S$100 cashback when they apply and spend S$800 within 60 days of card approval. This offer is exclusive to customers new to Standard Chartered Credit Cards. Terms and conditions apply."]]],

          ["What Caltex fuel benefit applies to the Standard Chartered Simply Cash Credit Card?",
           [["New Simply Cash Credit Cardholders receive six S$5 instant Caltex fuel discount vouchers. A minimum fuel spend of S$70 is required for each voucher use. Upon card approval, the discount vouchers are sent by push notification and can be accessed in the SC Mobile app. Existing Simply Cash Credit Cardholders enjoy an upfront discount of 17% on all fuel purchases."]]],

          ["What is the annual fee on the Standard Chartered Simply Cash Credit Card?",
           [["The annual fee on the Simply Cash Credit Card is waived for the first year. Thereafter, the annual fee is S$196.20 (including 9% GST)."]]],

          ["What finance charges apply to the Standard Chartered Simply Cash Credit Card?",
           [[FINANCE]]],

          ["What cash advance fee applies to the Standard Chartered Simply Cash Credit Card?",
           [[CASH_ADV]]],

          ["What late payment charge applies to the Standard Chartered Simply Cash Credit Card?",
           [[LATE]]],

          ["Who is eligible to apply for the Standard Chartered Simply Cash Credit Card?",
           ["Application eligibility for the Simply Cash Credit Card is:",
            B([["Singapore Citizens and Permanent Residents: aged 21 to 65 years, with a minimum annual income of S$30,000."],
               ["Foreigners: aged 21 to 65 years, with a minimum annual income of S$90,000, and must hold a Singapore Employment Pass."]])]],

          ["Is the EZ-Link facility still available on the Standard Chartered Simply Cash Credit Card?",
           [["The EZ-Link facility on the Simply Cash Credit Card was discontinued from 1 June 2024. If the Card was issued before December 2019 and stored value remains in the EZ-Link account or purse, the cardholder should visit any SimplyGo Ticket Office to obtain a refund."]]],
        ],
      },
      {
        name: "Smart Credit Card",
        url: U.smart,
        qas: [
          ["What cashback rates apply on the Standard Chartered Smart Credit Card?",
           ["From 6 December 2024, the Smart Credit Card earns cashback based on total card spend in the statement month:",
            B([["Card spend below S$800: 4% on bonus categories and 0.5% on other eligible spend. The 4% rate for spend below S$800 was valid only for a limited period from 6 December 2024 to 31 March 2025."],
               ["Card spend of S$800 or more: 8% on bonus categories and 0.5% on other eligible spend."],
               ["Card spend of S$1,500 or more: 10% on bonus categories and 1% on other eligible spend."]])]],

          ["Which merchant categories earn bonus cashback on the Standard Chartered Smart Credit Card?",
           ["A total of 20 merchants across three pillars are eligible to earn bonus cashback on the Smart Credit Card:",
            B([["Smart Dining: McDonald’s, Burger King, KFC, Subway, Toast Box, Ya Kun Kaya Toast, Starbucks, The Coffee Bean & Tea Leaf, Pizza Hut and Domino’s Pizza."],
               ["Smart Streaming: Netflix, Disney+, Spotify Premium, YouTube Premium, Amazon Prime, Viu, iQIYI and HBO Go."],
               ["Smart Transport: Bus and MRT rides via SimplyGo, and Electric Vehicle charging. All transactions under MCC 5552 for EV charging are included and valid only until 31 December 2025."]])]],

          ["What is earned on the Standard Chartered Smart Credit Card if the minimum S$800 monthly spend is not met?",
           [["Clients who charge less than S$800 to the Smart Credit Card in their statement month earned 4% cashback on bonus categories from 6 December 2024 to 31 March 2025. From 1 April 2025, a cardholder who does not meet the minimum S$800 monthly spend earns the base reward of 0.5% cashback on all eligible spend, including bonus categories."]]],

          ["Is there a cap on the cashback earned on the Standard Chartered Smart Credit Card?",
           [["There is no cap on cashback from 6 December 2024 onwards. Under the previous Smart Card Rewards Promotions terms and conditions there was a cap of 14,400 360° Rewards Points (the “Smart Cap”) per statement month. That cap was removed from 6 December 2024, so Smart Cardholders enjoy no cap on both bonus categories and any eligible spend."]]],

          ["Does the Standard Chartered Smart Credit Card earn cashback or Rewards Points?",
           ["All eligible spend on the Smart Credit Card, including bonus categories, earns 360° Rewards Points. The points earned per statement month are:",
            B([["Card spend below S$800: 12.8 points (or 4% cashback) on bonus categories and 1.6 points (or 0.5% cashback) on other eligible spend."],
               ["Card spend of S$800 or more: 25.6 points (or 8% cashback) on bonus categories and 1.6 points (or 0.5% cashback) on other eligible spend."],
               ["Card spend of S$1,500 or more: 32.0 points (or 10% cashback) on bonus categories and 3.2 points (or 1% cashback) on other eligible spend."]])]],

          ["How is cashback redeemed from Rewards Points on the Standard Chartered Smart Credit Card?",
           [["Cashback is redeemed through the SC Online Rewards page. Cashback is redeemed in blocks of S$10, and a minimum of 3,200 360° Rewards Points is required for each redemption of S$10 cashback."]]],

          ["When are Rewards Points credited on the Standard Chartered Smart Credit Card?",
           ["Crediting depends on the type of spend:",
            B([["All 360° Rewards Points for bonus categories earned within the statement month are credited 1 day after the statement cycle date. For example, if the statement cycle is on the 19th, the spend period runs from the 20th to the 19th of the following month."],
               ["All 360° Rewards Points for other eligible spend earned within the statement month are credited at the end of the statement cycle date. For example, if the statement cycle is on the 19th, those points are credited on the 19th."]])]],

          ["Does an existing Standard Chartered Smart Credit Card need to be replaced following the December 2024 benefit changes?",
           [["There is no need to replace an existing Smart Credit Card. Cardholders can continue using the card they hold."]]],

          ["When did the changes to the Standard Chartered Smart Credit Card benefits take effect?",
           [["The changes took effect from the cardholder’s next statement cycle after 6 December 2024. For example, if the statement cycle is on the 19th, the first statement month under the new Smart Card benefits ran from 20 December 2024 to 19 January 2025."]]],

          ["Can transactions still be converted to 3-month interest-free instalments on the Standard Chartered Smart Credit Card?",
           [["Transactions of a minimum of S$150 can still be converted to interest-free instalments on the Smart Credit Card. However, a processing fee is charged for instalment conversions from 13 January 2025. Any EasyPay instalments effected on the Smart Credit Card before 12 January 2025 are not impacted and no processing fee is charged for them."]]],

          ["What is the annual fee on the Standard Chartered Smart Credit Card?",
           ["The annual fee on the Smart Credit Card is S$91 (excluding 9% GST), applied as follows:",
            B([["New cardholders: the annual fee is waived for the first year, and charged for subsequent years."],
               ["Existing cardholders as at 6 December 2024: the annual fee is waived for the next 12 months, and from 6 December 2025 an annual fee of S$91 (excluding 9% GST) applies."]])]],

          ["What finance charges apply to the Standard Chartered Smart Credit Card?",
           [["From 6 December 2024, all Smart Cardholders have a standardised Effective Interest Rate (EIR) of 27.9% per annum. Previously, Smart Cardholders were assigned an EIR of 23.9%, 27.9% or 29.9% per annum based on the Bank’s assessment of the cardholder’s credit profile. Finance charges apply only if payment is not made in full by the due date. The change took effect from the cardholder’s next statement cycle after 6 December 2024."]]],

          ["What cash advance fee applies to the Standard Chartered Smart Credit Card?",
           [["The cash advance fee is waived for Smart Credit Cardholders."]]],

          ["What late payment charge applies to the Standard Chartered Smart Credit Card?",
           [[LATE]]],

          ["Who is eligible to apply for the Standard Chartered Smart Credit Card?",
           ["Application eligibility for the Smart Credit Card is:",
            B([["Singapore Citizens and Permanent Residents: aged 21 to 65 years, with a minimum annual income of S$30,000."],
               ["Foreigners: aged 21 to 65 years, with a minimum annual income of S$90,000, and must hold a Singapore Employment Pass."]])]],
        ],
      },
    ],
  },

  // ------------------------------------------------------------------------
  {
    name: "Miles and Travel Cards",
    subs: [
      {
        name: "Visa Infinite Credit Card",
        url: U.visaInfinite,
        qas: [
          ["How are miles maximised on the Standard Chartered Visa Infinite Credit Card?",
           ["Charging a minimum of S$2,000 on eligible spends in a statement cycle earns up to 3 miles for every S$1 spent:",
            B([["Local spend: 3.5 Rewards Points (1.4 miles) per S$1."],
               ["Foreign currency spend: 7.5 Rewards Points (3 miles) per S$1."],
               ["All spends below S$2,000 in a statement cycle earn 2.5 points, or 1 mile, for every S$1 spent."]])]],

          ["How are 360° Rewards Points converted to KrisFlyer miles on the Standard Chartered Visa Infinite Credit Card?",
           [["360° Rewards Points are redeemed for KrisFlyer miles through the 360° Rewards platform. Log in with SC Online credentials, then select “Travel” followed by “Singapore Airlines – KrisFlyer”. Points are redeemed in blocks of 25,000 Rewards Points for 10,000 airline miles, and a transfer fee of S$27.25 (including GST) applies to each redemption transaction. The redemption platform is at ", ["sc.com/sg/rewards-programmes/360-rewards-redemption", L.rewardsRedeem], "."]]],

          ["How is complimentary Priority Pass membership obtained on the Standard Chartered Visa Infinite Credit Card?",
           [["To apply for complimentary Priority Pass membership, activate the principal Standard Chartered Visa Infinite Credit Card and then SMS SCVI<space>PP<space>last 4 digits of the credit card number to 77272, using the mobile number registered with the Bank. For example: SCVI PP 1234. The card provides six complimentary visits each year to Priority Pass lounges worldwide, at over 1,700 locations."]]],

          ["Is the annual fee on the Standard Chartered Visa Infinite Credit Card waivable?",
           [["The annual fee charged on the Standard Chartered Visa Infinite Credit Card is strictly non-waivable. The annual fee is S$594 (including 8% GST), adjusted to S$599.50 (including 9% GST) from 1 January 2024."]]],

          ["How is the complimentary travel insurance on the Standard Chartered Visa Infinite Credit Card activated?",
           [["Complimentary travel insurance worth S$1,000,000 applies when the cardholder charges the full travel fare to the Standard Chartered Visa Infinite Credit Card before going abroad. Terms and conditions apply."]]],

          ["Can the Income Tax Payment Facility be used with Standard Chartered credit cards other than the Visa Infinite Card?",
           [["The Income Tax Payment Facility is only available for the Standard Chartered Visa Infinite Card."]]],

          ["How does the Income Tax Payment Facility work on the Standard Chartered Visa Infinite Credit Card?",
           ["Charging income tax to the Visa Infinite Credit Card earns miles, with a preferential one-time processing fee of 1.9% on the tax payable amount. The process is:",
            B([["Complete the application form within 14 business days before the tax payable due date."],
               ["The approved amount is credited to the chosen bank account, and processing fees are charged to the Visa Infinite Credit Card."],
               ["The outcome of the application is notified by SMS."],
               ["For faster processing, ensure the Visa Infinite Credit Card has sufficient credit limit to cater for the tax payable amount and processing fees during the application."]])]],

          ["What renewal bonus applies on the Standard Chartered Visa Infinite Credit Card?",
           [["A renewal bonus of 50,000 Rewards Points, which can be redeemed for 20,000 KrisFlyer Miles, is awarded upon the Standard Chartered Visa Infinite Credit Card anniversary and payment of the annual fee of S$550 (excluding GST)."]]],

          ["Who is eligible to apply for the Standard Chartered Visa Infinite Credit Card?",
           ["Application eligibility for the Visa Infinite Credit Card is age 21 to 65 years old, with minimum annual income as follows:",
            B([["Priority or Private Banking customers who are Singapore Citizens or Permanent Residents: S$30,000."],
               ["Priority or Private Banking customers who are foreigners with Employment Passes: S$60,000."],
               ["All other customers: S$150,000."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Visa Infinite Credit Card?",
           ["The charges on the Visa Infinite Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "Journey Credit Card",
        url: U.journey,
        qas: [
          ["What miles rates apply on the Standard Chartered Journey Credit Card?",
           ["The Journey Credit Card earns miles at the following rates:",
            B([["3 miles per S$1 on selected bonus categories: transportation, food deliveries and online groceries."],
               ["2 miles per S$1 on overseas spends, with no cap."],
               ["1.2 miles per S$1 on local spends, with no cap."]])]],

          ["Which bonus spend categories earn 3 miles per dollar on the Standard Chartered Journey Credit Card?",
           ["Bonus categories are online transactions denominated in Singapore dollars in the following Merchant Category Codes (MCCs) and merchants:",
            B([["Transport merchants with MCCs 4111, 4121, 4411 and 4789. Examples include Grab, Gojek, Cabcharge Asia, Tada, Ryde Technologies, Royal Caribbean Cruises and Easybook.com."],
               ["Grocery and food stores with MCCs 5411, 5462, 5499 and 5921. Examples include NTUC FairPrice Online, NTUC FairPrice app, Lazada Redmart, Watson’s Singapore and Nespresso Singapore."],
               ["Food deliveries with MCCs 5811, 5812 and 5814. Examples include Foodpanda, Deliveroo, McDonald’s, KFC, Pizza Hut, Dominos Pizza, Chilli Api Catering and Qi Ji Catering."]])]],

          ["How are online transactions defined for the bonus spend categories on the Standard Chartered Journey Credit Card?",
           [["Online transactions are card-not-present (CNP) transactions identified via merchant category codes (MCC). These category codes are assigned by Visa, merchants and acquiring banks, not by Standard Chartered Bank (Singapore) Limited. If a transaction does not fall under the assigned category codes, it will not qualify for the 3 miles bonus. Apple Pay, Samsung Pay, Google Pay and Garmin Pay mobile wallet transactions are not CNP transactions. The qualifying MCC list is in the ", ["Journey Credit Card Terms & Conditions", L.journeyTnc], "."]]],

          ["What happens when bonus category spend exceeds S$1,000 in a statement cycle on the Standard Chartered Journey Credit Card?",
           [["Total spend in the bonus categories is capped at S$1,000 for each statement month, equivalent to a maximum of 3,000 miles (7,500 points). Any additional spend beyond S$1,000 earns 1.2 miles per dollar (3 points)."]]],

          ["When are the 3 miles per dollar for bonus spend credited on the Standard Chartered Journey Credit Card?",
           ["Miles on the Journey Credit Card are credited as 360° Rewards Points:",
            B([["Base 360° Rewards Points are credited when the eligible transaction is posted."],
               ["Bonus 360° Rewards Points are aggregated and credited on the statement cycle date."]])]],

          ["Is there a cap on miles earned for non-bonus category transactions on the Standard Chartered Journey Credit Card?",
           [["There is no cap. Cardholders continue to earn 1.2 miles for every S$1 spent locally and 2.0 miles for every S$1 spent overseas on the Journey Credit Card."]]],

          ["How are 360° Rewards Points converted to KrisFlyer miles on the Standard Chartered Journey Credit Card?",
           [["360° Rewards Points are redeemed for KrisFlyer miles through the 360° Rewards platform. Log in with SC Online credentials, then select “Travel” followed by “Singapore Airlines – KrisFlyer”. Points are redeemed in blocks of 25,000 Rewards Points for 10,000 airline miles, and a transfer fee of S$27.25 (including GST) applies to each redemption transaction."]]],

          ["Which transactions are not eligible to earn miles on the Standard Chartered Journey Credit Card?",
           ["Transactions that are not eligible to earn miles on the Journey Credit Card include:",
            B([["Cash advances or Credit Card Funds Transfers."],
               ["Payments to charitable or political organisations."]])]],

          ["What is the annual fee on the Standard Chartered Journey Credit Card?",
           [["The annual fee for the principal Journey Credit Card is S$196.20 (including 9% GST) from 1 January 2024. If the “FEE WAIVED” option was selected when applying for the card, an annual fee auto-waiver applies for the first year. For annual fees in the second and subsequent years, 10,000 miles are awarded upon renewal of the Card and payment of the annual fee, credited by the end of the following month."]]],

          ["Who is eligible to apply for the Standard Chartered Journey Credit Card?",
           ["Application eligibility for the Journey Credit Card is:",
            B([["Singapore Citizens and Permanent Residents: aged 21 to 65 years, with a minimum annual income of S$30,000."],
               ["Foreigners: aged 21 to 65 years, with a minimum annual income of S$90,000, and must hold a Singapore Employment Pass."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Journey Credit Card?",
           ["The charges on the Journey Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "Beyond Credit Card",
        url: U.beyond,
        qas: [
          ["What miles rates apply on the Standard Chartered Beyond Credit Card?",
           ["Earn rates on the Beyond Credit Card depend on the cardholder’s banking relationship:",
            B([["Beyond Credit Cardholder: 1.5 miles per S$1 of local spend and 3 miles per S$1 of foreign spend."],
               ["Priority Banking plus Beyond Credit Cardholder: 2 miles per S$1 of local spend and 3.5 miles per S$1 of foreign spend."],
               ["Priority Private plus Beyond Credit Cardholder: 2 miles per S$1 of local spend, and on foreign spend 8 miles per S$1 on dining and 4 miles per S$1 on others."]])]],

          ["How much is the annual fee on the Standard Chartered Beyond Credit Card?",
           [["There is an annual fee of S$1,500 (excluding GST) for the principal Beyond Credit Card, and no additional fees for supplementary Beyond Credit Cards. The annual card membership fee for the Beyond Credit Card is not eligible for waiver."]]],

          ["How is a banking relationship status checked for the Standard Chartered Beyond Credit Card?",
           [["Log in to the SC Mobile App and go to Services. The current banking relationship status is shown under the cardholder’s name."]]],

          ["Can a Personal Banking client apply for the Standard Chartered Beyond Credit Card?",
           [["Yes, a Personal Banking client may apply for the Beyond Credit Card. Application eligibility is age 21 years and above with a minimum annual income of S$200,000, and foreigners must hold a Singapore Employment Pass."]]],

          ["How is a banking relationship upgraded to Priority Banking or Priority Private for the Standard Chartered Beyond Credit Card?",
           [["To upgrade a banking relationship, top up a Standard Chartered current or savings account to a minimum account balance of S$200,000 for Priority Banking or S$1,500,000 for Priority Private. Upon reaching the relevant minimum account balance threshold, log in to the SC Mobile App and go to Services > View all > Other Banking Services > Priority Banking Sign-up Request."]]],

          ["How is the Standard Chartered Beyond Credit Card rewards earn rate affected when banking relationship status changes mid-cycle?",
           [["When banking relationship status is updated, the newly applicable rewards earn rate is only applied in the next statement cycle and subsequently awarded in the following card statement cycle. For example, for an existing Priority Banking client with a card statement dated 15 Jan whose relationship was upgraded to Priority Private on 5 Jan, the Priority Banking rewards earn rate still applies to all posted transactions from 5 to 15 Jan. In the next statement cycle ending 15 Feb the new Priority Private earn rate applies, and this is awarded in the 15 Mar statement cycle."]]],

          ["Why do the Rewards Points on a Standard Chartered Beyond Credit Card statement differ from the expected amount?",
           [["Only the Base Rewards Points are credited in the current card statement, while Additional Rewards Points are credited in the next card statement. Additional Rewards Points refer to all Rewards Points granted in excess of what a non-Priority Banking client receives for the equivalent transactions made."]]],

          ["Is there a cap on Rewards Points or miles earned on the Standard Chartered Beyond Credit Card?",
           [["The Beyond Credit Card offers miles, rewarded in Rewards Points, with no earning cap."]]],

          ["How are miles earned viewed and converted on the Standard Chartered Beyond Credit Card?",
           [["360° Rewards Points are earned on eligible spending and can be redeemed in blocks of 25,000 Rewards Points for 10,000 airline miles. The number of Rewards Points earned for the latest card statement month is shown at the bottom of the latest card statement. Points are converted to miles on the ", ["360° Rewards Redemption Platform", L.rewardsRedeem], ". Terms and conditions apply."]]],

          ["What travel privileges are included with the Standard Chartered Beyond Credit Card?",
           ["Travel privileges on the Beyond Credit Card include:",
            B([["Complimentary Business Class upgrades. From 22 November 2025 for a limited time, a complimentary Business Class upgrade applies on renowned airlines when one Business Class ticket and one Premium Economy ticket are purchased."],
               ["Unlimited Priority Pass airport lounge visits for the cardholder and Supplementary Cardholders, plus an additional six complimentary visits for guests per year."],
               ["Complimentary travel insurance coverage of up to USD 500,000, including travel delays, missed connections, loss of baggage, medical expenses and medical evacuation, when trips are booked with the card."],
               ["Hotel elite loyalty status: GHA Discovery Titanium tier and Wyndham Hotels and Resorts Diamond tier membership, plus ALL Accor+ Explorer membership."],
               ["Airport limousine rides."]])]],

          ["What dining privileges are included with the Standard Chartered Beyond Credit Card?",
           ["Dining privileges on the Beyond Credit Card include:",
            B([["A complimentary birthday meal at a designated Michelin-starred restaurant in Singapore for the principal cardholder, with 15% off for accompanying loved ones. The redemption code is available in the SC Mobile app two weeks before the cardholder’s birth month."],
               ["The Luxe Plate: 15% off at over 25 renowned restaurants in Singapore, plus an invitation-only event featuring Michelin-starred local and international chefs."],
               ["Priority reservation access at over 1,000 of the world’s finest dining venues across more than 140 cities."]])]],

          ["What lifestyle and protection privileges are included with the Standard Chartered Beyond Credit Card?",
           ["Lifestyle and protection privileges on the Beyond Credit Card include:",
            B([["Four complimentary green fees at 74 premium golf clubs across Southeast Asia, plus complimentary golf lessons at Golf Performance 360 Golf Academy, Singapore."],
               ["Purchase Protection covering theft and accidental damage for up to 180 days on online purchases made with the card."],
               ["Complimentary enrolment into the ID Theft Protect Programme to monitor and receive alerts on identity theft."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Beyond Credit Card?",
           ["The charges on the Beyond Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
    ],
  },

  // ------------------------------------------------------------------------
  {
    name: "Rewards Points Cards",
    subs: [
      {
        name: "Rewards+ Credit Card",
        url: U.rewardsPlus,
        qas: [
          ["What Rewards Points rates apply on the Standard Chartered Rewards+ Credit Card?",
           ["The Rewards+ Credit Card earns up to 10x Rewards Points on foreign currency spend for overseas retail, dining and travel, and up to 5x Rewards Points on dining transactions in Singapore dollars. In detail:",
            B([["Qualifying Dining Transactions: 1 Reward Point plus up to 4 additional Rewards Points for every S$1 spent."],
               ["Foreign currency transactions: 1 Reward Point plus up to 9 additional Rewards Points for every S$1 spent."],
               ["The additional Rewards Points are subject to the Rewards+ Cap."]])]],

          ["How is a Foreign Currency transaction defined on the Standard Chartered Rewards+ Credit Card?",
           [["Foreign Currency spend is defined as non-SGD denominated transactions. Where a retail merchant overseas offers a choice to pay either in the local currency of the country being visited or in SGD, choosing the local currency is deemed a Foreign Currency transaction and choosing SGD is deemed an SGD transaction."]]],

          ["What are “Qualifying Dining Transactions” on the Standard Chartered Rewards+ Credit Card?",
           ["Qualifying Dining Transactions are dining spends charged to the Rewards+ Card at dining establishments in Singapore classified under the Merchant Category Codes (MCC) of:",
            B([["Caterers."],
               ["Restaurants and eating places."],
               ["Pubs and bars."],
               ["Fast Food Restaurants."],
               ["Purchases from establishments not classified under those MCCs, including but not limited to local hotel restaurants (including wedding banquets), bakeries and supermarkets, are not considered Qualifying Dining Transactions."]])]],

          ["Is dining at hotels a Qualifying Dining Transaction on the Standard Chartered Rewards+ Credit Card?",
           [["Dining at hotels may or may not qualify, because it depends on the MCC category of the restaurant within the hotel. In general, hotel restaurants are typically categorised under the “Lodging” MCC. As “Lodging” does not fall within the Qualifying Dining Transaction MCCs, such spend is not an eligible Qualifying Dining Transaction, and only 1 Reward Point is awarded for every S$1 charged to the Rewards+ Credit Card."]]],

          ["How many Rewards Points are earned on a Qualifying Dining Transaction made in foreign currency on the Standard Chartered Rewards+ Credit Card?",
           [["A Qualifying Dining Transaction made in foreign currency is recognised as a foreign currency transaction, so the cardholder receives the higher rate of 9 additional Rewards Points for every S$1 spent."]]],

          ["How does the Rewards+ Cap of 20,000 Rewards Points work on the Standard Chartered Rewards+ Credit Card?",
           ["The Rewards+ cap of 20,000 Rewards Points applies only to the additional Rewards Points awarded on Qualifying Dining Transactions and foreign currency transactions, namely:",
            B([["Additional four 360° Reward Points for every S$1 spent on Qualifying Dining Transactions; or"],
               ["Additional nine 360° Reward Points for every S$1 spent in foreign currency, meaning spends that are not made in Singapore Dollars, on Qualifying Rewards+ Card Transactions."]])]],

          ["Is the Rewards+ Cap based on the calendar year on the Standard Chartered Rewards+ Credit Card?",
           [["No. The Rewards+ Cap of 20,000 Rewards Points applies to a full year from the approval date of the Rewards+ Card, and is reset every anniversary of the approval date."]]],

          ["Can more than 20,000 Rewards Points be earned on the Standard Chartered Rewards+ Credit Card?",
           [["Yes. The Rewards+ Cap of 20,000 Rewards Points applies only to the additional Rewards Points awarded on Qualifying Dining Transactions and foreign currency transactions. Once the cap is reached, the cardholder continues to earn 1 Reward Point per S$1 spent on the card."]]],

          ["What are Qualifying Rewards+ Card transactions on the Standard Chartered Rewards+ Credit Card?",
           [["Qualifying Rewards+ Card transactions are retail transactions, with the exception of excluded transactions. The excluded transactions are listed at ", ["sc.com/sg/rewardsplustncs", L.rewardsPlusTnc], "."]]],

          ["Can a supplementary cardholder use the Rewards+ Cathay Movie Promotion on the Standard Chartered Rewards+ Credit Card?",
           [["Yes, the Rewards+ Cathay Movie Promotion is available to both main and supplementary Rewards+ cardholders."]]],

          ["How is the complimentary Cathay Cineplex Platinum Movie Suites Ticket redeemed on the Standard Chartered Rewards+ Credit Card?",
           [["Purchase the Cathay Cineplex Platinum Movie Suites Ticket over the counter and present the Rewards+ Credit Card for payment. The complimentary ticket is provided on purchase of an equivalent ticket."]]],

          ["Is there a limit on complimentary Cathay Cineplex Platinum Movie Suites Tickets on the Standard Chartered Rewards+ Credit Card?",
           [["Yes. Each Rewards+ cardholder is limited to a maximum of 5 daily redemptions, subject to a maximum of 1,500 redemptions for the promotion."]]],

          ["What is the annual fee on the Standard Chartered Rewards+ Credit Card?",
           [["The annual fee on the Rewards+ Credit Card is S$196.20 (including 9% GST) from 1 January 2024. This annual fee is waived for the first year."]]],

          ["What complimentary travel insurance applies on the Standard Chartered Rewards+ Credit Card?",
           [["The Rewards+ Credit Card provides complimentary travel medical insurance coverage of up to S$500,000. The cardholder must charge the full travel fare to the card before going abroad. Terms and conditions apply."]]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Rewards+ Credit Card?",
           ["The charges on the Rewards+ Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "NUS Alumni Platinum Card",
        url: U.nus,
        qas: [
          ["What are the benefits of the Standard Chartered NUS Alumni Platinum Card?",
           [["The NUS Alumni Platinum Card lets the cardholder earn Rewards Points on their spend and contribute to their alma mater. Standard Chartered contributes a percentage of the cardholder’s spend towards supporting NUS students in bursaries and awards. The card also includes a 24-hour Visa Platinum Concierge for travel arrangements, reservations, shopping and gift services."]]],

          ["Is the Standard Chartered NUS Alumni Platinum Card only available to NUS graduates?",
           [["Yes, the NUS Alumni Platinum Credit Card is exclusive to NUS graduates."]]],

          ["Is a minimum spend required to qualify for Rewards Points on the Standard Chartered NUS Alumni Platinum Card?",
           [["No, there is no minimum spend required to qualify for the Rewards Points on the NUS Alumni Platinum Card."]]],

          ["What is the annual fee on the Standard Chartered NUS Alumni Platinum Card?",
           [["The annual fee on the NUS Alumni Platinum Card is waived for the first year. Thereafter, the annual fee is S$196.20 (including 9% GST) from 1 January 2024."]]],

          ["Who is eligible to apply for the Standard Chartered NUS Alumni Platinum Card?",
           ["Application eligibility for the NUS Alumni Platinum Card is age 21 to 65 years old, with minimum annual income and pass requirements as follows:",
            B([["Singapore Citizens and Permanent Residents: S$30,000."],
               ["Foreigners with Employment Passes: S$90,000."],
               ["Foreigners must hold P1, P2 or Q type Singapore Employment Passes. Q Pass holders must have a minimum of one year validity remaining on their passes, and the Employment Pass must have at least 6 months’ validity."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered NUS Alumni Platinum Card?",
           ["The charges on the NUS Alumni Platinum Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "Platinum Visa / Mastercard Credit Card",
        url: U.platinum,
        qas: [
          ["Is the Standard Chartered Platinum Visa / Mastercard Credit Card still available?",
           [["No. Standard Chartered Singapore no longer offers the Platinum Visa / Mastercard Credit Card. Other card options are listed at ", ["sc.com/sg/credit-cards", L.cards], ", including the Smart Credit Card and the Unlimited Cashback Card."]]],

          ["What Rewards Points rate applied on the Standard Chartered Platinum Visa / Mastercard Credit Card?",
           [["The Platinum Visa / Mastercard Credit Card earned 1 Rewards Point for every S$1 charged to the credit card."]]],

          ["What was the annual fee on the Standard Chartered Platinum Visa / Mastercard Credit Card?",
           [["The annual fee on the Platinum Visa / Mastercard Credit Card is S$196.20 (including 9% GST) from 1 January 2024. This annual fee was waived for the first three years."]]],

          ["What was the eligibility for the Standard Chartered Platinum Visa / Mastercard Credit Card?",
           ["Application eligibility for the Platinum Visa / Mastercard Credit Card was age 21 to 65 years old, with minimum annual income and pass requirements as follows:",
            B([["Singapore Citizens and Permanent Residents: S$30,000."],
               ["Foreigners with Employment Passes: S$90,000."],
               ["Foreigners must hold P1, P2 or Q type Singapore Employment Passes. Q Pass holders must have a minimum of one year validity remaining on their passes, and the Employment Pass must have at least 6 months’ validity."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Platinum Visa / Mastercard Credit Card?",
           ["The charges on the Platinum Visa / Mastercard Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
    ],
  },

  // ------------------------------------------------------------------------
  {
    name: "Priority Banking and Partner Cards",
    subs: [
      {
        name: "Priority Banking Visa Infinite Credit Card",
        url: U.priorityVI,
        qas: [
          ["How are 360° Rewards Points earned on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           ["The 360° Rewards programme rewards the total banking relationship, not only card spend. Points are earned as follows:",
            B([["Priority Banking Visa Infinite Credit Card: spend S$1 and earn 1 air mile (2.5 points)."],
               ["Current or savings account: 10 points per month for every S$10,000 average balance."],
               ["Fixed deposits: 10 points per month for every S$10,000 average balance."],
               ["Investments: 5 points per month for every S$20,000 average Asset Under Management."],
               ["Mortgage loans: 10 points per month for every S$10,000 outstanding balance."],
               ["A spend of S$500 per month is required to enjoy the total banking relationship 360° Rewards Points, which are credited to the card every month."]])]],

          ["Do all transactions earn Rewards Points on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           [["Not all transactions qualify for Rewards Points. An eligible retail transaction must be charged to earn Rewards Points. The list of transactions that do not qualify is set out in the card’s full terms and conditions at ", ["av.sc.com/sg/content/docs/sg-pbvi-tnc-full.pdf", L.pbviTnc], "."]]],

          ["Do Rewards Points expire on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           [["There is no expiry date on Rewards Points earned on the Standard Chartered Priority Banking Visa Infinite Credit Card."]]],

          ["What can 360° Rewards Points be redeemed for on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           ["Highlights of the rewards redemption catalogue are:",
            B([["40,350 points: S$100 Cash Credit into the card account."],
               ["25,000 points: 10,000 Singapore Airlines KrisFlyer miles."]])]],

          ["How is complimentary Priority Pass membership obtained on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           [["To apply for complimentary Priority Pass membership, activate the principal Standard Chartered Priority Banking Visa Infinite Credit Card and then SMS PBVI<space>PP<space>last 4 digits of the credit card number to 77272, using the mobile number registered with the Bank. For example: PBVI PP 1234."]]],

          ["How is the complimentary travel medical insurance activated on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           [["Complimentary Travel Insurance worth $50,000 each for the cardholder and their spouse applies when the card is charged at least once while overseas. Terms and conditions apply."]]],

          ["Can the annual fee on the Standard Chartered Priority Banking Visa Infinite Credit Card be waived?",
           [["A waiver of the annual fee charged to the Standard Chartered Priority Banking Visa Infinite Card is available to fully funded Priority Banking clients."]]],

          ["What is the annual fee on the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           ["The annual fees on the Priority Banking Visa Infinite Credit Card from 1 January 2024 are:",
            B([["Principal card: S$327 (including 9% GST). This annual fee is waived for the first year."],
               ["Supplementary card: S$163.50 (including 9% GST). This annual fee is waived for the first 5 cards."]])]],

          ["Who is eligible to apply for the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           ["Application eligibility for the Priority Banking Visa Infinite Credit Card is age 21 to 65 years old, with minimum annual income as follows, for Priority or Private Banking customers:",
            B([["Singapore Citizens and Permanent Residents: S$30,000."],
               ["Foreigners with Employment Passes: S$60,000."],
               ["Foreigners must hold P1, P2 or Q type Singapore Employment Passes."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Priority Banking Visa Infinite Credit Card?",
           ["The charges on the Priority Banking Visa Infinite Credit Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "Prudential Platinum Card",
        url: U.pruPlatinum,
        qas: [
          ["What Rewards Points rates apply on the Standard Chartered Prudential Platinum Card?",
           ["The Prudential Platinum Card earns Rewards Points at the following rates:",
            B([["10X Rewards Points for every S$1 spent in foreign currency on Qualifying Foreign Currency Transactions."],
               ["5X Rewards Points for every S$1 spent in local currency on Qualifying Dining Transactions."],
               ["1X Reward Point for every S$1 spent on all other eligible retail transactions, including insurance transactions."]])]],

          ["Is a minimum spend required to earn the additional Rewards Points on the Standard Chartered Prudential Platinum Card?",
           [["No, there is no minimum spend amount required to qualify for the additional Rewards Points under the Prudential Card Rewards Points Promotion on the Prudential Platinum Card."]]],

          ["How does the Rewards Cap of 20,000 Rewards Points work on the Standard Chartered Prudential Platinum Card?",
           ["The Rewards Cap of 20,000 Rewards Points applies only to the additional 360° Rewards Points awarded on Qualifying Dining Transactions and Qualifying Foreign Currency Transactions. It is in addition to the total Rewards Points earned for other qualifying spends, which earn 1X Rewards Point for every S$1 or its equivalent. The additional Rewards Points are:",
            B([["For all Qualifying Dining Transactions, an additional 4X Reward Points for every S$1 spent in Singapore Dollars."],
               ["For every Qualifying Foreign Currency Transaction, an additional 9X Reward Points for every S$1 spent in foreign currency, meaning spends that are not made in Singapore Dollars."]])]],

          ["Is the Rewards Cap on the Standard Chartered Prudential Platinum Card based on the calendar year?",
           [["No. The Rewards Cap of 20,000 Rewards Points applies to a full year from the approval date of the Prudential Platinum Credit Card and is reset every anniversary of the approval date."]]],

          ["Can more than 20,000 Rewards Points be earned on the Standard Chartered Prudential Platinum Card?",
           [["Yes. The Rewards Cap of 20,000 Rewards Points applies only to the additional Rewards Points awarded on Qualifying Dining Transactions and Qualifying Foreign Currency Transactions. Other qualifying spend continues to earn 1X Rewards Point for every S$1 spent."]]],

          ["Can Rewards Points be earned when paying Prudential premiums with the Standard Chartered Prudential Platinum Card?",
           [["Yes, Rewards Points are earned when the Prudential Platinum Card is used to pay Prudential premiums."]]],

          ["How are recurring Prudential premium payments set up with the Standard Chartered Prudential Platinum Card?",
           ["Recurring payments can be applied for and paid using the Prudential Platinum Credit Card on the following websites:",
            B([[["PRUaccess", L.pruAccess]],
               [["iPay", L.iPay]]])]],

          ["How are Rewards Points used to offset Prudential premiums with the Standard Chartered Prudential Platinum Card?",
           [["Points can be converted to offset premium payments for Eligible Prudential Policies. With effect from 1 June 2023, each redemption must have a minimum conversion of 3,000 points to offset S$9.60 in Prudential insurance premium payments."]]],

          ["What is the annual fee on the Standard Chartered Prudential Platinum Card?",
           [["The annual fee on the Prudential Platinum Card is waived for the first year. Thereafter, the annual fee is S$196.20 (including 9% GST) from 1 January 2024."]]],

          ["Who is eligible to apply for the Standard Chartered Prudential Platinum Card?",
           ["Application eligibility for the Prudential Platinum Card is age 21 to 65 years old, with minimum annual income and pass requirements as follows:",
            B([["Singapore Citizens and Permanent Residents: S$30,000."],
               ["Foreigners with Employment Passes: S$90,000."],
               ["Foreigners must hold P1, P2 or Q type Singapore Employment Passes. Q Pass holders must have a minimum of one year validity remaining on their passes, and the Employment Pass must have at least 6 months’ validity."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Prudential Platinum Card?",
           ["The charges on the Prudential Platinum Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
      {
        name: "Prudential Visa Signature Card",
        url: U.pruSignature,
        qas: [
          ["Who can apply for the Standard Chartered Prudential Visa Signature Card?",
           [["The policyholder must be an Ascend by Prudential customer to be eligible for the Prudential Visa Signature Card."]]],

          ["What Rewards Points rates apply on the Standard Chartered Prudential Visa Signature Card?",
           ["The Prudential Visa Signature Card earns Rewards Points at the following rates:",
            B([["10X Rewards Points for every S$1 spent in foreign currency on Qualifying Foreign Currency Transactions."],
               ["5X Rewards Points for every S$1 spent in local currency on Qualifying Dining Transactions."],
               ["1.5X Reward Points for every S$1 spent on all other eligible retail transactions, including insurance transactions."]])]],

          ["Is a minimum spend required to earn the additional Rewards Points on the Standard Chartered Prudential Visa Signature Card?",
           [["No, there is no minimum spend amount required to qualify for the additional Rewards Points under the Prudential Card Rewards Points Promotion on the Prudential Visa Signature Card."]]],

          ["How does the Rewards Cap of 50,000 Rewards Points work on the Standard Chartered Prudential Visa Signature Card?",
           ["The Rewards Cap of 50,000 Rewards Points applies only to the additional Rewards Points awarded on Qualifying Dining Transactions and Qualifying Foreign Currency Transactions. It is in addition to the total Rewards Points earned for other qualifying spends, which earn 1.5X Rewards Points for every S$1 or its equivalent. The additional Rewards Points are:",
            B([["For all Qualifying Dining Transactions, an additional 3.5X Reward Points for every S$1 spent in Singapore Dollars."],
               ["For every Qualifying Foreign Currency Transaction, an additional 8.5X Reward Points for every S$1 spent in foreign currency, meaning spends that are not made in Singapore Dollars."]])]],

          ["Is the Rewards Cap on the Standard Chartered Prudential Visa Signature Card based on the calendar year?",
           [["No. The Rewards Cap of 50,000 Rewards Points applies to a full year from the approval date of the Prudential Visa Signature Credit Card and is reset every anniversary of the approval date."]]],

          ["Can more than 50,000 Rewards Points be earned on the Standard Chartered Prudential Visa Signature Card?",
           [["Yes. The Rewards Cap of 50,000 Rewards Points applies only to the additional Rewards Points awarded on Qualifying Dining Transactions and Qualifying Foreign Currency Transactions. Other qualifying spend continues to earn 1.5X Rewards Points for every S$1 spent."]]],

          ["Can Rewards Points be earned when paying Prudential premiums with the Standard Chartered Prudential Visa Signature Card?",
           [["Yes, Rewards Points are earned when the Prudential Visa Signature Card is used to pay Prudential premiums."]]],

          ["How are recurring Prudential premium payments set up with the Standard Chartered Prudential Visa Signature Card?",
           ["Recurring payments can be applied for and paid using the Prudential Visa Signature Credit Card on the following websites:",
            B([[["PRUaccess", L.pruAccess]],
               [["iPay", L.iPay]]])]],

          ["How are Rewards Points used to offset Prudential premiums with the Standard Chartered Prudential Visa Signature Card?",
           [["Points can be redeemed to offset insurance premiums for Eligible Prudential Policies. With effect from 1 June 2023, each redemption must have a minimum conversion of 3,000 points to offset S$9.60 in Prudential insurance premium payments."]]],

          ["What complimentary travel insurance applies on the Standard Chartered Prudential Visa Signature Card?",
           [["The Prudential Visa Signature Card provides complimentary travel insurance coverage of up to S$500,000. The cardholder must charge the full travel fare to the card before going abroad. Terms and conditions apply."]]],

          ["What is the annual fee on the Standard Chartered Prudential Visa Signature Card?",
           [["The annual fee on the Prudential Visa Signature Card is waived for the first year. Thereafter, the annual fee is S$196.20 (including 9% GST) from 1 January 2024."]]],

          ["Who is eligible to apply for the Standard Chartered Prudential Visa Signature Card?",
           ["Application eligibility for the Prudential Visa Signature Card requires the policyholder to be an Ascend by Prudential customer, aged 21 to 65 years old, with minimum annual income and pass requirements as follows:",
            B([["Singapore Citizens and Permanent Residents: S$30,000."],
               ["Foreigners with Employment Passes: S$90,000."],
               ["Foreigners must hold P1, P2 or Q type Singapore Employment Passes. Q Pass holders must have a minimum of one year validity remaining on their passes, and the Employment Pass must have at least 6 months’ validity."]])]],

          ["What finance charges, cash advance fee and late payment charge apply to the Standard Chartered Prudential Visa Signature Card?",
           ["The charges on the Prudential Visa Signature Card are:",
            B([[FINANCE],
               [CASH_ADV],
               [LATE]])]],
        ],
      },
    ],
  },

  // ------------------------------------------------------------------------
  {
    name: "Applying for a Credit Card",
    subs: [
      {
        name: "Supporting Documents",
        url: U.simplyCash,
        qas: [
          ["What documents does an existing Standard Chartered credit cardholder need to apply for another credit card?",
           [["No documents are required from existing Standard Chartered Credit Cardholders. However, if there has been a recent change to the applicant’s income, updated documents must be submitted for review during the application process."]]],

          ["What documents does a Singaporean or Permanent Resident salaried employee need to apply for a Standard Chartered credit card?",
           ["SingPass holders applying with MyInfo need no documents. Other salaried employees who are Singapore Citizens or Permanent Residents need:",
            B([["Copy of NRIC (front and back)."],
               ["Latest computerised payslip, or latest 6 months’ CPF Transaction History Statement."],
               ["To be considered for a higher credit limit, the latest Income Tax Notice of Assessment (NOA) in addition to the above."]])]],

          ["What documents does a self-employed Singaporean or Permanent Resident need to apply for a Standard Chartered credit card?",
           [["A self-employed Singapore Citizen or Permanent Resident must provide a copy of a 2 year income tax notice of assessment, and must have been in business for a minimum of 2 years."]]],

          ["What documents does a commission earner who is a Singaporean or Permanent Resident need to apply for a Standard Chartered credit card?",
           ["A commission earner who is a Singapore Citizen or Permanent Resident must provide any one of the following:",
            B([["Latest 3 months payslip plus latest income tax notice of assessment."],
               ["Latest 6 months CPF Transaction History Statement."],
               ["Latest 3 months payslip."],
               ["Copy of 2 year income tax notice of assessment, for a 100% commission earner."]])]],

          ["What identity and address documents does a foreigner need to apply for a Standard Chartered credit card?",
           ["A foreign applicant must prepare:",
            B([["Copy of Passport with at least 6 months’ validity, including the page with address displayed where applicable."],
               ["Copy of the applicant’s Employment Pass with at least 6 months’ validity."],
               ["Any ONE of: latest utility bill, rates or tax bill; latest bank or credit card statement (e-Statements are accepted); rental agreement showing the address; latest mobile phone statement or pay-TV statement; letter from employer stating current address; or a government-issued document stating current address, for example from IRAS, CPF or ICA."]])]],

          ["What income documents does a foreigner need to apply for a Standard Chartered credit card?",
           ["Income documents required from a foreign applicant depend on employment type:",
            B([["Salaried employees and partial commission-based earners: latest computerised payslip. To be considered for a higher loan amount, the latest Income Tax Notice of Assessment (NOA) in addition to the above."],
               ["100% commission-based earners: copy of NRIC (front and back), plus either the latest Income Tax Notice of Assessment (NOA) or the latest 3 months of Commission Statement from the same employer."],
               ["Self-employed: latest Income Tax Notice of Assessment (NOA)."]])]],

          ["Which documents are no longer included in the Standard Chartered credit card welcome pack?",
           ["As part of Standard Chartered’s efforts to go green, the following are no longer included within the credit card welcome pack:",
            B([["The ABS Consumer Guide for Credit Cards, which contains information on understanding credit cards and credit card usage."],
               ["The Important Information Document for Standard Chartered Credit Cards, which contains information on all relevant interest, fees and charges. Fees and charges are also published in the ", ["Standard Chartered Singapore Pricing Guide", L.pricingGuide], "."]])]],

          ["How is a free credit report obtained after a Standard Chartered credit card is approved?",
           [["If a credit card is newly approved, a free credit report can be obtained within 30 calendar days of card approval via the credit bureau website at ", ["creditbureau.com.sg", L.creditBureau], " or at the bureau’s registered office. The card carrier that comes with the credit card and the NRIC are required. The free credit report is not available for renewed or replaced cards."]]],
        ],
      },
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Credit Cards FAQ",
  docTitle: "SCB Bank_Credit Cards FAQ",
  description: "Standard Chartered Singapore Credit Cards FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Credit_Cards_FAQ.docx", buf);
  console.log("written", buf.length);
});
