const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

const S = "https://www.sc.com/sg/save/";
const U = {
  hub: S,
  basic: S + "savings-accounts/basic-bank-account/",
  esaver: S + "savings-accounts/esaver/",
  esaverKids: S + "savings-accounts/esaver-kids/",
  fcysaver: S + "savings-accounts/fcysaver/",
  firstsaver: S + "savings-accounts/firstsaver/",
  jumpstart: S + "savings-accounts/jumpstart/",
  jumpstartStory: S + "savings-accounts/jumpstart/stories/getting-your-first-credit-card-here-are-some-must-knows/",
  myway: S + "savings-accounts/myway/",
  usdsaver: S + "savings-accounts/usdsaver/",
  bonussaver: S + "current-accounts/bonussaver/",
  chequesave: S + "current-accounts/cheque-and-save-account/",
  supersalary: S + "current-accounts/supersalary/",
  wealthsaver: S + "current-accounts/wealth-saver/",
  tdFcy: S + "time-deposits/foreign-currency-time-deposits/",
  tdSgd: S + "time-deposits/singapore-dollar-time-deposit/",
  tdSustainable: S + "time-deposits/sustainable-time-deposit/",
  mca: S + "mca/",
};

const L = {
  moneyLock: "https://www.sc.com/sg/bank-with-us/money-lock/",
  pricingGuide: "https://www.sc.com/sg/pricing-guide/",
  remit: "https://www.sc.com/sg/bank-with-us/remittance/",
  simplyCash: "https://www.sc.com/sg/credit-cards/simply-cash-credit-card/",
};

// Supporting-document boilerplate shared across the deposit account pages.
const DOCS_SGP = "Singapore citizens and permanent residents must provide their NRIC, front and back.";
const DOCS_FOREIGN_INTRO = "Foreign applicants must provide a passport and an Employment Pass, plus any ONE of the following as proof of address:";
const DOCS_FOREIGN = [
  ["Utility bill within the last 3 months."],
  ["Bank statement within the last 3 months."],
  ["Letter from the human resource department of the current employer stating the address."],
  ["Government-issued documents stating the address, for example from IRAS, CPF or ICA."],
  ["Telecommunications bill within the last 3 months."],
  ["Tenancy agreement."],
];

const CATEGORIES = [
  // ======================================================= SAVINGS ACCOUNTS
  {
    name: "Savings Accounts",
    subs: [
      { name: "Basic Bank Account", url: U.basic, qas: [
        ["Who is eligible for the Standard Chartered Basic Bank Account?",
         ["Eligibility for the Basic Bank Account is:",
          B([["Nationality: Singapore citizen under the Government’s Public Assistance Scheme or Special Grant Scheme."],
             ["Age requirement: minimum 15 years old."],
             ["Initial deposit: $0."]])]],

        ["What documents are required to open a Standard Chartered Basic Bank Account?",
         ["Singapore citizens must provide the following mandatory supporting documents:",
          B([["NRIC, front and back."],
             ["Government’s Public Assistance Scheme or Special Grant Scheme Card and/or Letter."]])]],

        ["What charges apply to the Standard Chartered Basic Bank Account?",
         ["The key charges on the Basic Bank Account are:",
          B([["Minimum initial deposit: S$20, waived for beneficiaries under the MSF Public Assistance Scheme or Special Grant Scheme."],
             ["Minimum average daily balance: S$0."],
             ["Fall-below fee: S$0."],
             ["Monthly service charge: S$2 regardless of balance, waived for beneficiaries under the MSF Public Assistance Scheme or Special Grant Scheme."],
             ["Early account closure fee within 6 months: S$30."],
             ["Passbook replacement fee: S$30."]])]],

        ["What features come with the Standard Chartered Basic Bank Account?",
         ["The Basic Bank Account includes the following features:",
          B([["Cash withdrawals at Standard Chartered’s network of branches."],
             ["NETS payments island wide with the ATM card."],
             ["Free withdrawals at more than 200 atm5 machines in Singapore, identified by the atm5 logo."],
             ["24-hour global access to cash through the Cirrus and Maestro networks using the ATM card."],
             ["Standing instructions for regular payments."],
             ["Online payments to over 360 local billing organisations registered with the Bank, paid with a Standard Chartered credit card or directly from the account."]])]],
      ]},

      { name: "e$aver Account", url: U.esaver, qas: [
        ["What is the Standard Chartered e$aver Savings Account?",
         ["The e$aver Savings Account is a high-yield savings account with the following features:",
          B([["No lock-in period, so high interest rates are earned on savings without any lock-in."],
             ["Instant access to funds anytime via SC Mobile or Online Banking."],
             ["Overseas fund transfers at S$0 charges with ", ["SC Remit", L.remit], "."]])]],

        ["What is the e$aver Savings Account Bonus Interest Promotion for August to September 2026?",
         ["From 1 August 2026 to 30 September 2026, up to 1.60% p.a. interest is earned on an eligible deposit balance of S$2 million with no lock-in period, when incremental fresh funds are brought in compared to the July 2026 average daily balance. The rates on the eligible incremental balance are:",
          B([["Qualifying accounts where the primary account holder is a customer with wealth holding: 0.05% p.a. prevailing plus 1.55% p.a. bonus, totalling 1.60% p.a."],
             ["Qualifying accounts where the primary account holder is a customer without wealth holding: 0.05% p.a. prevailing plus 1.35% p.a. bonus, totalling 1.40% p.a."]])]],

        ["What are “fresh funds” for the e$aver Savings Account promotion?",
         [["Fresh funds mean funds that do not originate from any existing account with Standard Chartered, or funds that are not withdrawn and re-deposited within 30 days of the promotion period."]]],

        ["What is a “customer with wealth holdings” for the e$aver Savings Account?",
         [["Customers with wealth holdings are defined as customers who hold either an Eligible Investment Product or an Eligible Insurance Product purchased through the Bank at the end of each calendar month during the promotion period. An Eligible Investment Product refers to Unit Trusts, Bonds, Structured Notes or Equities, excluding foreign currency exchange transactions and investments using the Central Provident Fund Investment Scheme or Supplementary Retirement Scheme. The eligible Bonus Interest Rate is assessed based on wealth holdings at the end of each calendar month during the promotion period."]]],

        ["Who is eligible to open a Standard Chartered e$aver Account?",
         [["A person is eligible to open an e$aver savings account as a primary account holder if they are 18 years old and above. For joint account applications, the joint account holder can be between the ages of 12 and 18 years old. The account is open to Singapore citizens, Singapore permanent residents and foreigners."]]],

        ["What is average daily balance on a Standard Chartered account?",
         [["Average daily balance (ADB) is the sum of end-of-day account balances of all days in a particular calendar month, divided by the number of days in that month."]]],

        ["When is bonus interest credited to a Standard Chartered e$aver Account?",
         [["Bonus interest is credited to the account by the 25th of the following month at the latest. For example, for the month of August 2026, bonus interest is credited by 25 September 2026 at the latest."]]],

        ["What charges apply to the Standard Chartered e$aver Account?",
         ["The key charges on the e$aver Account are:",
          B([["Minimum initial deposit: S$0."],
             ["Minimum average daily balance: S$1,000."],
             ["Fall-below fee: S$5 per month."],
             ["Early account closure fee within 6 months: S$0."],
             ["Counter transaction: S$5 per cash withdrawal."]])]],

        ["What documents are required to open a Standard Chartered e$aver Account?",
         [DOCS_SGP + " " + DOCS_FOREIGN_INTRO, B(DOCS_FOREIGN)]],
      ]},

      { name: "e$aver Kids Account", url: U.esaverKids, qas: [
        ["What are the benefits of the Standard Chartered e$aver Kids Account?",
         ["The e$aver Kids Account has the following benefits:",
          B([["No minimum initial deposit to open an account for a child."],
             ["No minimum balance, so any balance can be maintained."],
             ["No lock-in period, giving the flexibility to withdraw funds anytime."],
             ["No monthly fees, so savings build up faster."]])]],

        ["Who is eligible for a Standard Chartered e$aver Kids Account?",
         ["Eligibility for the e$aver Kids Account is:",
          B([["The Trustee, meaning the parent, must be 21 years old and above."],
             ["The Beneficiary, meaning the child, must be below 18 years old."],
             ["Singaporeans, permanent residents or foreigners may apply."]])]],

        ["What is required to open a Standard Chartered e$aver Kids Account?",
         ["The account opening requirements are:",
          B([["The Trustee must complete the e$aver Kids Trust Account Opening Form."],
             ["The Trustee must sign the Declaration of Trust."],
             ["The Trustee must submit the NRIC or passport of the Trustee, of which a copy is retained."],
             ["The Trustee must submit the birth certificate of the beneficiary, of which a copy is retained."],
             ["The Trustee must complete a GIRO or Standing Instruction Application Form for monthly deposits of at least $50 into the account, which must be from a bank account in the Trustee’s name."]])]],

        ["What charges apply to the Standard Chartered e$aver Kids Account?",
         ["The key charges on the e$aver Kids Account are:",
          B([["Minimum initial deposit: S$0."],
             ["Fall-below fee: S$0."],
             ["Minimum average daily balance: S$0."],
             ["Early account closure fee within 6 months: S$0."]])]],
      ]},

      { name: "FCY$aver Account", url: U.fcysaver, qas: [
        ["What is the Standard Chartered FCY$aver Account?",
         [["The FCY$aver Account is a foreign currency savings account that can be opened with as little as USD 2,000 or equivalent. It is available in USD, EUR, GBP, AUD, NZD, CHF, CAD, HKD and JPY, and can be opened instantly using MyInfo."]]],

        ["What interest rate applies to the Standard Chartered FCY$aver Account?",
         [["The prevailing interest rate for FCY$aver balances is 0.00% p.a."]]],

        ["What are the minimum average daily balances and fall-below fees on the Standard Chartered FCY$aver Account?",
         ["The minimum average daily balance and monthly fall-below fee by currency are:",
          B([["AUD 2,000 with a fall-below fee of AUD 5. CAD 2,000 with CAD 5. CHF 2,000 with CHF 5. NZD 2,000 with NZD 5. USD 2,000 with USD 5."],
             ["EUR 1,000 with a fall-below fee of EUR 3. GBP 1,000 with GBP 2."],
             ["HKD 15,000 with a fall-below fee of HKD 30. JPY 200,000 with JPY 500."]])]],

        ["What are the deposit and closure charges on the Standard Chartered FCY$aver Account?",
         ["The charges on the FCY$aver Account are:",
          B([["Minimum initial deposit: USD 2,000 or its equivalent."],
             ["Early account closure fee within 6 months: USD 20."]])]],

        ["Who is eligible to open a Standard Chartered FCY$aver Account, and what documents are required?",
         ["The FCY$aver Account is open to Singapore citizens, Singapore permanent residents and foreigners aged 18 years and above. " + DOCS_SGP + " " + DOCS_FOREIGN_INTRO,
          B(DOCS_FOREIGN)]],
      ]},

      { name: "First$aver Account", url: U.firstsaver, qas: [
        ["What is the Standard Chartered First$aver Account?",
         [["First$aver is a joint savings account created to empower youths with financial autonomy while ensuring parents or guardians have oversight of their child’s financial journey. Youths aged 13 and above have access to a debit card, mobile banking and digital payments, and parents can view their child’s transactions via SC Mobile or Online Banking. For accounts where the child is under 13 years old, only parents or guardians can transact with the account."]]],

        ["What interest rate does the Standard Chartered First$aver Account pay?",
         [["First$aver offers a prevailing interest rate of 1.00% p.a. This interest rate applies only to the first S$50,000 deposited in the account. No fees or minimum balance are required."]]],

        ["What cashback does the Standard Chartered First$aver debit card give?",
         [["Youths aged 13 and above enjoy 1% cashback on qualifying transactions on their First$aver debit card, up to a maximum of S$60 per month."]]],

        ["Who is eligible for a Standard Chartered First$aver Account?",
         ["Eligibility for the First$aver Account is:",
          B([["Parent: Singaporean or Singapore Permanent Resident, 26 years old and above."],
             ["Child: Singaporean, Singapore Permanent Resident, or foreigner born in Singapore, below 18 years old."],
             ["Application is via online channels with MyInfo only, through Online Banking, SC Mobile or the Standard Chartered website."]])]],

        ["How is a Standard Chartered First$aver Account applied for?",
         ["A parent aged 26 years and above applies online as follows:",
          B([["Step 1: Select “Apply Now”."],
             ["Step 2: Pre-fill the application via SingPass (MyInfo)."],
             ["Step 3: Select the child the account is being created for."],
             ["Step 4: Verify the parent’s and the child’s details."],
             ["Step 5: Submit the application, and the account is opened instantly."]])]],

        ["How many Standard Chartered First$aver Accounts can be held?",
         [["Each child can only hold one First$aver account. Parents can hold a separate account with each child."]]],

        ["Are there fees and charges on the Standard Chartered First$aver Account?",
         [["There are no fees and charges applicable for the First$aver Account."]]],

        ["What are the spend and transfer limits on the Standard Chartered First$aver Account?",
         ["The limits on the First$aver Account are:",
          B([["The daily spend limit (POS limit) on the debit card defaults to S$100, which the youth can adjust up to a maximum of S$250 via SC Mobile."],
             ["The daily ATM withdrawal limit on the debit card defaults to S$100."],
             ["The daily transfer limit on the account defaults to S$100."]])]],

        ["Does a parent receive a debit card on the Standard Chartered First$aver Account?",
         [["No. The First$aver Debit Card is only issued to youths 13 years old and above."]]],

        ["What can youths do in SC Mobile on a Standard Chartered First$aver Account?",
         ["Youths are provided a simplified view on SC Mobile and can only perform the following:",
          B([["View transactions and balances."],
             ["Local Transfers, PayNow, Scan & Pay and Add Payees."],
             ["View account statements."],
             ["Services relating to their debit card."]])]],

        ["Who has Online Banking access on a Standard Chartered First$aver Account?",
         [["Youths can only access mobile banking via SC Mobile and do not have Online Banking access. Only parents have full access to both SC Mobile and Online Banking."]]],

        ["Can Money Lock be used on a Standard Chartered First$aver Account?",
         [["Only parents can lock and unlock the funds in the First$aver account. Youths can view how much money is locked away. Money Lock is an anti-scam security feature that helps protect funds from scams and unauthorised transfers or withdrawals, and it can also prevent incidental overspending by a child. Parents start Money Lock via SC Mobile by logging in, selecting the account, indicating the amount and submitting the request. To unlock funds, a request must be submitted at any Standard Chartered Singapore branch, and locked funds are released in full on the same day. Further details are at ", ["sc.com/sg/bank-with-us/money-lock", L.moneyLock], "."]]],

        ["What features and access apply by age on a Standard Chartered First$aver Account?",
         ["Access differs by role and age:",
          B([["Child under 13 years old: no debit card, no SC Mobile, no Online Banking and no Money Lock."],
             ["Youth 13 years old and above: debit card yes, SC Mobile yes, Online Banking no, and Money Lock view of the locked balance only."],
             ["Parents: no debit card, SC Mobile yes, Online Banking yes, and Money Lock with the ability to view, lock and unlock the balance."]])]],

        ["What is a savings account, as explained for First$aver youths?",
         [["A savings account is an account at a bank that holds a person’s money. The Bank pays money, called interest, when the money is kept in the account."]]],

        ["What is interest in a savings account, as explained for First$aver youths?",
         [["Interest is the money paid by the Bank for keeping money with them. Interest is paid as a percentage of the money put in the account at a certain rate, commonly known as the interest rate per annum. For example, with $1,000 in a savings account at an interest rate of 1% per annum, 1% of $1,000 is $10, so at the end of one year the account holds $1,010. The more money in the savings account, the more interest is earned."]]],

        ["What is a bank transaction, as explained for First$aver youths?",
         [["A bank transaction is any money that has moved in and out of a savings account. Examples include paying with a debit card, sending money via PayNow, or a parent depositing money into the account. All these transactions are recorded in the account and can be viewed via the SC Mobile app."]]],

        ["What is FAST, as explained for First$aver youths?",
         [["Fast And Secure Transfers (FAST) is an electronic funds transfer service that allows money to be sent from one bank account to another bank account in Singapore using an account number. To send money via FAST, a payee must first be added under the Local payee type via SC Mobile, entering the Payee Bank, Payee Name and Payee Account Number. Once a payee is added successfully, money is sent by selecting “Local Transfer” under the Pay & Transfer tab in SC Mobile."]]],

        ["What is PayNow, as explained for First$aver youths?",
         [["PayNow is a simpler and faster way to send and receive money instantly from one bank account to another using a mobile number, Singapore NRIC/FIN number or UEN (Unique Entity Number). A friend’s bank account details do not need to be added to the account when transferring money via PayNow."]]],

        ["What is Scan & Pay, as explained for First$aver youths?",
         [["Scan & Pay allows a retail shop’s QR code to be scanned to make a payment. The correct amount should be keyed in before approving the payment."]]],

        ["What is a debit card, as explained for First$aver youths?",
         [["A debit card can be used to pay for things instead of cash, and the amount is deducted from the First$aver Account. It can be used at any shop that accepts contactless payments with a tap, and cash can be withdrawn at any Standard Chartered ATM in Singapore using a 5-digit PIN. It can also be used to buy things online, which requires the name, card number, card expiration date and CVV. The card and its information should be kept safe and not shared with anyone, as they can be misused by other people to make transactions without the cardholder’s knowledge."]]],

        ["How can a Standard Chartered First$aver debit card be used?",
         ["The First$aver debit card can be used in the following ways:",
          B([["Withdraw cash at any Standard Chartered ATM or participating bank’s ATM in Singapore, or at an ATM overseas, by entering the PIN and selecting the Savings account."],
             ["Deposit cash at a Cash Deposit Machine in Singapore by entering the PIN and selecting the Savings account."],
             ["Pay via NETS by entering the PIN at the NETS terminal."],
             ["Pay via Visa or Mastercard in Singapore and overseas by inserting the card at the payment terminal and signing for the purchase. Overseas usage must be activated before use."],
             ["Make contactless payments via Visa or Mastercard by tapping the card at the payment terminal, or by adding the debit card to a mobile device and tapping without the physical card."]])]],

        ["What is a CVV, as explained for First$aver youths?",
         [["CVV stands for Card Verification Value, a 3-digit unique number printed on the back of the card. It is used to authorise online payments and helps protect the cardholder if the debit card number is stolen. If someone obtains the card number but not the CVV, it is harder for them to make fraudulent transactions, so the CVV must not be disclosed to anyone."]]],

        ["What is a NETS transaction, as explained for First$aver youths?",
         [["A NETS transaction is a way to make payments at shops that accept NETS, using the First$aver debit card. A 5-digit PIN must be entered to approve the NETS transaction."]]],

        ["What is payWave or Tap & Go, as explained for First$aver youths?",
         [["payWave and Tap & Go are ways to make contactless payments using a card, by waving or tapping the card at the payment terminal in shops which accept debit cards."]]],

        ["How is a Standard Chartered debit card activated on the SC Mobile app?",
         [["Tap the Services tab located at the bottom right corner, click the “Digital Services” menu, select “View All”, then select “Card Management” followed by “Debit/ATM Card Activation & PIN Set”."]]],

        ["How is a lost or stolen Standard Chartered debit card reported on SC Mobile?",
         ["A lost or stolen debit card is reported as follows:",
          B([["Step 1: Log in to Online Banking or SC Mobile."],
             ["Step 2: Go to “Help & Services”."],
             ["Step 3: Select “Report Lost/Stolen Card”."],
             ["Step 4: Follow the on-screen instructions."]])]],
      ]},

      { name: "JumpStart Account", url: U.jumpstart, qas: [
        ["What is the Standard Chartered JumpStart Account?",
         ["JumpStart is a simple, high-interest savings account for young adults and students, offering the following:",
          B([["A Base interest rate of 0.50% p.a. for account balances up to S$50,000."],
             ["A Step-Up interest rate of 1.00% p.a. on deposit balances up to S$50,000 when the account holder invests."],
             ["1% cashback on eligible debit card spends."],
             ["No lock-in period, no minimum deposit, no minimum spends and no requirement for salary crediting."],
             ["A Cashback debit card with no monthly or annual fees."]])]],

        ["What are the requirements to open a Standard Chartered JumpStart Account?",
         [["To open a JumpStart Account, the applicant must be between 18 and 26 years old at the time that they apply for the account."]]],

        ["Must a Standard Chartered JumpStart Account be closed at age 26?",
         [["There is no requirement to close a JumpStart Account on turning 26. The account can continue to be maintained beyond the account holder’s 26th birthday."]]],

        ["How many Standard Chartered JumpStart Accounts can be opened, and can they be joint accounts?",
         [["Only one JumpStart Account can be opened. A JumpStart account can only be opened in the account holder’s own name and cannot be a joint-named account."]]],

        ["How is interest calculated on a Standard Chartered JumpStart Account?",
         [["Prevailing Base interest is received on JumpStart account balances and credited into the account at the end of each month. The account earns 0.50% p.a. on the first $50,000 of account balance, and 0.10% on any incremental balances above S$50,000. There is no cap to the amount of interest that can be received on the balances in a JumpStart account."]]],

        ["How does interest appear on a Standard Chartered JumpStart Account statement?",
         [["Interest is reflected in the eStatement, Online Banking and Mobile Banking transaction history as “CR INTEREST”, and is credited into the JumpStart Account at the end of each month. Step-Up interest is reflected as “CR STEP UP INTEREST”."]]],

        ["Is interest paid if a Standard Chartered JumpStart Account is closed before the crediting date?",
         [["If the JumpStart Account is closed before the end of the month, prevailing interest is paid up to the day before account closure."]]],

        ["Are there fees on the Standard Chartered JumpStart Account?",
         ["There are no fees on the JumpStart Account:",
          B([["There is no minimum deposit requirement."],
             ["There is no monthly fee."],
             ["There is no fall-below fee."],
             ["There is no service or annual fee for the Cashback debit card."]])]],

        ["When is the Standard Chartered JumpStart Cashback debit card received?",
         [["The Cashback debit card is received 5 to 7 working days from the opening date of the JumpStart Savings account."]]],

        ["What cashback applies on the Standard Chartered JumpStart Cashback debit card?",
         [["1% cashback is earned on qualifying transactions on the debit card. Monthly cashback is capped at S$60 per account and is calculated at the end of each calendar month based on transaction posting date. Cashback for debit card spends is credited on the first working day of the following month."]]],

        ["How is Step-Up Interest earned on a Standard Chartered JumpStart Account?",
         ["Step-Up Interest is earned by completing any one of the following in a calendar month:",
          B([["Successfully subscribe and settle one or more Eligible Unit Trusts through any investment account held as a primary account holder."],
             ["Successfully set up a Regular Savings Plan (RSP) and settle the monthly investment through any investment account held as a primary account holder."],
             ["Successfully put through at least one buy order for Equities via SC Online Trading and settle the trades through any investment account held as a primary account holder."]])]],

        ["What is an Eligible Unit Trust for Standard Chartered JumpStart Step-Up Interest?",
         [["An Eligible Unit Trust means a unit trust distributed by the Bank, but does not include exchange traded funds, switching transactions, or any investments made through an Overdraft Facility, the Central Provident Fund Investment Scheme or the Supplementary Retirement Scheme."]]],

        ["What are Equities for Standard Chartered JumpStart Step-Up Interest?",
         [["Equities comprise all shares that are open for trade on the Standard Chartered Online Trading platform, but do not include delisted or suspended shares, or shares that are in certificate form."]]],

        ["Is there a minimum investment amount to earn Standard Chartered JumpStart Step-Up Interest?",
         [["There is no minimum investment amount requirement to receive Step-Up Interest. However, most Unit Trusts require a minimum investment of S$1,000, with subsequent amounts as low as S$100, and the monthly minimum to set up a Regular Savings Plan is $100 but may be higher depending on each Unit Trust. For Equities, the minimum investment value in terms of number of shares that can be traded online is determined by the respective exchanges."]]],

        ["Can more than 1.00% p.a. Step-Up Interest be earned in a month on a Standard Chartered JumpStart Account?",
         [["No. The maximum Step-Up Interest received per month is at the rate of 1% p.a., even if Unit Trusts, a Regular Savings Plan and Equities are all invested in during the same month."]]],

        ["When is Step-Up Interest credited to a Standard Chartered JumpStart Account?",
         [["Step-Up Interest is received for each month in which a Wealth Quest is completed. If the Wealth Quest is fulfilled in a given calendar month, Step-Up Interest is credited into the JumpStart account the following calendar month."]]],

        ["What Step-Up Interest would a S$50,000 Standard Chartered JumpStart balance earn?",
         ["Standard Chartered gives the following illustrations, each based on a JumpStart average daily balance of S$50,000 in January:",
          B([["Performing one or more buy transactions on the SC Online Trading platform in January: estimated Step-Up interest credited in February is S$41.66."],
             ["Performing one buy transaction on SC Online Trading and one on the Online Unit Trust platform in January: estimated Step-Up interest credited in February is S$41.66."],
             ["Setting up a Regular Savings Plan to invest $100 each month for the next 12 months in January: estimated Step-Up interest of S$41.66 per month for the next 12 months."]])]],

        ["What documents are required to open a Standard Chartered JumpStart Account at a branch?",
         ["To apply for a JumpStart account at a branch:",
          B([["Singapore citizens and permanent residents: original Singapore NRIC."],
             ["Foreigners residing in Singapore: passport, a valid student, long term or employment pass, and any one of the last 3 months’ utility bill, last 3 months’ bank statement, telecommunications bill, tenancy agreement, or government-issued documents stating the address such as from IRAS, CPF or ICA."]])]],
      ]},

      { name: "MyWay Savings Account", url: U.myway, qas: [
        ["What is the Standard Chartered MyWay Savings Account?",
         [["MyWay is a savings account tailored for those 55 years old and above. It offers attractive interest rates of up to 1.50% p.a. on the first S$5 million of balances, and a complimentary Digital Scam Protection Insurance which protects against digital theft of funds from the MyWay Savings Account, capped at S$50,000."]]],

        ["Who is eligible to apply for a Standard Chartered MyWay Savings Account?",
         [["MyWay Savings Account is exclusively for clients who are at least 55 years old at the time of account opening. For joint applications, the main applicant must be at least 55 years old while the joint applicant must be at least 18 years old. Foreigners must hold P1, P2 or Q type Singapore Employment Passes, and Q Pass holders must have a minimum of one year validity remaining on their passes."]]],

        ["What interest rates apply to the Standard Chartered MyWay Savings Account?",
         ["Effective from 1 May 2026, the prevailing interest rates by account balance are:",
          B([["First S$50,000: 0.05% p.a."],
             ["Next S$200,000: 0.30% p.a."],
             ["Next S$1,250,000: 0.50% p.a."],
             ["Next S$3,500,000: 1.50% p.a."],
             ["Above S$5,000,000: 0.05% p.a. The maximum effective interest rate is 1.19% p.a. for a deposit balance of S$5,000,000."]])]],

        ["What interest rates applied to the Standard Chartered MyWay Savings Account before 1 May 2026?",
         ["Before 1 May 2026, the prevailing interest rates by account balance were:",
          B([["First S$50,000: 0.05% p.a."],
             ["Next S$200,000: 0.30% p.a."],
             ["Next S$1,250,000: 0.60% p.a."],
             ["Next S$3,500,000: 2.00% p.a."],
             ["Above S$5,000,000: 0.05% p.a. The maximum effective interest rate was 1.56% p.a. for a deposit balance of S$5,000,000."]])]],

        ["How is interest calculated on a Standard Chartered MyWay Savings Account?",
         [["The interest on a MyWay Savings Account is calculated at the end of each day and credited to the MyWay Savings Account at the end of each month."]]],

        ["Who is eligible for the Standard Chartered MyWay Savings Account Bonus Interest Promotion?",
         [["The MyWay Savings Account Bonus Interest Promotion is open to all primary account holders of a MyWay Savings Account who bring in a minimum of S$200,000 of Fresh Funds into their MyWay Savings Account during the Promotion Period, such that the month-end deposit balance in the month the deposit is made, the Participating Month, is at least S$200,000 higher than the month-end deposit balance of the previous month."]]],

        ["What does the MyWay Digital Scam Protection Insurance cover?",
         [["Digital Scam Protection Insurance is a complimentary insurance offered to MyWay Savings Account holders aged 55 years old and above. It covers loss of funds up to S$50,000 from the MyWay Savings Account due to unauthorised electronic transfer of money arising from phishing or email spoofing. Exclusions apply, including but not limited to losses from authorised payment scams such as impersonation scams, love scams or investment scams. The insurance is provided and underwritten by MSIG Insurance (Singapore) Pte. Ltd."]]],

        ["How is a MyWay Digital Scam Protection Insurance claim made?",
         [["Once any digital theft of funds is discovered, Standard Chartered must be notified and a police report lodged within 7 days from the discovery of the theft of funds. A claim form must then be filed with MSIG within 30 days from the discovery of the theft of funds. The Bank is not responsible for the processing, assessment or settlement of any claims submitted under this insurance."]]],

        ["Until when does the MyWay Digital Scam Protection Insurance run?",
         [["The complimentary Digital Scam Protection Insurance has been extended until 16 March 2027. There will be a revision to the MyWay Savings Account’s Digital Scam Protection Insurance effective 1 September 2026."]]],

        ["Which security features protect a Standard Chartered MyWay Savings Account?",
         ["Two security features protect the MyWay Savings Account:",
          B([["Kill Switch, a self-service function that quickly suspends accounts if the account holder suspects their account has been compromised."],
             ["Money Lock, an anti-scam security feature available on SC Mobile that helps protect funds from scams and unauthorised transfers or withdrawals. Money that is not needed for everyday access is locked from current or savings accounts and continues to earn interest at current rates. Further details are at ", ["sc.com/sg/bank-with-us/money-lock", L.moneyLock], "."]])]],

        ["What other benefits come with a Standard Chartered MyWay Savings Account?",
         ["The MyWay Savings Account includes the following additional benefits:",
          B([["A Priority Banking upgrade with exclusive rewards on a minimum funding of S$200,000 in the MyWay Savings Account."],
             ["Preferential rates for health screening at AsiaMedic Wellness Assessment Centre Pte Ltd and other The Good Life offers with the MyWay Debit Card."],
             ["SC Shop and Earn, where Priority clients enjoy higher cashback at over 250 online merchants on top of their banking rewards."]])]],

        ["What charges apply to the Standard Chartered MyWay Savings Account?",
         ["The key charges on the MyWay Savings Account are:",
          B([["Minimum initial deposit: S$0."],
             ["Fall-below fee: S$10 per month if the average daily balance for the month falls below S$50,000."],
             ["Monthly service charge: S$0."],
             ["Early account closure fee within 6 months: S$0."],
             ["Debit card annual fee: S$0."]])]],
      ]},

      { name: "USD$aver Account", url: U.usdsaver, qas: [
        ["What is a USD savings account, and who should open one?",
         ["A USD savings account is a savings account that allows the holder to hold, transact and deposit in US Dollars. It suits:",
          B([["Frequent travellers who travel to the United States or other countries where the US Dollar is widely accepted."],
             ["Expats and foreign employees who earn income in US Dollars or need to send money home in US Dollars."],
             ["Investors looking to diversify their currency holdings and hedge against currency fluctuations."],
             ["Businesses engaged in international trade that make or receive transactions in US Dollars."],
             ["Parents and students managing funds for tuition payments and living expenses related to studying abroad in the US or other countries that use the US Dollar."]])]],

        ["What are the benefits of the Standard Chartered USD$aver Account?",
         ["The USD$aver Account offers the following benefits:",
          B([["Attractive high interest on USD deposits with no lock-in period."],
             ["The flexibility to withdraw funds anytime, anywhere."],
             ["2% cashback on all Mastercard transactions with the USD High Debit card, which has no annual fee. Exclusions apply and cashback is capped at USD100."],
             ["Savings on foreign currency conversion on USD denominated transactions."],
             ["Easy access to Online Banking and Mobile Banking to manage money and complete transactions anywhere in the world at any time."]])]],

        ["What prevailing interest rates apply to the Standard Chartered USD$aver Account?",
         ["The prevailing interest rates by balance tier are:",
          B([["First USD50,000: 0.05% p.a."],
             ["Next USD150,000: 0.50% p.a."],
             ["Next USD1,300,000: 1.00% p.a."],
             ["Next USD1,500,000: 2.50% p.a."],
             ["Account balances exceeding USD3,000,000 earn a prevailing interest of 0.05% p.a. The maximum effective interest rate on the prevailing rate is 1.71% p.a. for a deposit balance of USD3,000,000."]])]],

        ["What is the Standard Chartered USD$aver Top Up Promotion for August 2026?",
         ["The Top Up Promotion offers an additional 1.50% p.a. Bonus Interest on fresh funds deposits, on these conditions:",
          B([["Top up a minimum of USD200K in fresh funds by 31 August 2026."],
             ["Maintain the fresh funds in the account until 31 October 2026."],
             ["Earn 1.50% p.a. Bonus Interest on the incremental fresh funds for 3 months from August to October 2026."],
             ["Bonus Interest applies to the incremental average daily balance of each month during the Bonus Interest earning period compared to the 31 July 2026 balance, capped at USD3,000,000."],
             ["Fresh funds means funds that do not originate from any existing account with the Bank, or funds that are not withdrawn and re-deposited within 30 days of the Promotion Period."]])]],

        ["Is there a minimum deposit on the Standard Chartered USD$aver Account?",
         [["The USD$aver Account has a minimum average daily balance requirement of USD 10,000. If the average daily balance falls below the minimum amount for any particular month, there is a fall-below fee of USD10 per month. The minimum initial deposit is US$1."]]],

        ["What charges apply to the Standard Chartered USD$aver Account?",
         ["The key charges on the USD$aver Account are:",
          B([["Minimum initial deposit: US$1."],
             ["Fall-below fee: US$10 per month if the average daily balance for the month falls below US$10,000."],
             ["Monthly service charge: US$0."],
             ["Unarranged overdraft: Prime +5%, minimum US$5."],
             ["Early account closure fee within 6 months: US$20."],
             ["Debit card annual fee: US$0."]])]],

        ["Can Singaporeans, Permanent Residents and foreigners open a Standard Chartered USD$aver Account?",
         [["Yes. Singaporean and Permanent Resident applicants must be 18 years and above. Eligible non-Singaporean and non-Permanent Resident applicants must hold valid employment passes and must also be 18 years and above."]]],

        ["Can a Standard Chartered USD$aver Account be opened online, and can it be a joint account?",
         [["Yes. The USD$aver Account can be opened online through the Standard Chartered Singapore website. It can be opened as a personal account or as a joint account."]]],

        ["How many Standard Chartered USD$aver Accounts can be held?",
         [["A maximum of two USD$aver accounts can be held at any one time with Standard Chartered Singapore. However, each USD$aver account must have different account operating authority details."]]],

        ["How does the Standard Chartered USD$aver Account differ from a fixed deposit?",
         [["The USD$aver Account differs from a fixed deposit in that there is no lock-in period on the USD$aver account."]]],

        ["How are funds withdrawn from a Standard Chartered USD$aver Account?",
         [["Cash can be withdrawn from the USD$aver Account at an ATM using the USD High debit card, or at Standard Chartered branches. Applicable fees may apply."]]],
      ]},
    ],
  },

  // ======================================================= CURRENT ACCOUNTS
  {
    name: "Current Accounts",
    subs: [
      { name: "Bonus$aver Account", url: U.bonussaver, qas: [
        ["How is the maximum interest earned on a Standard Chartered Bonus$aver Account?",
         ["Up to 5.85% p.a. interest can be unlocked on the first S$100,000 of Average Daily Balance. Two types of interest are received every month, prevailing interest at 0.05% p.a. on the entire deposit balance, plus bonus interest on the first S$100,000 when these transactions are performed:",
          B([["Unlock 0.90% p.a. for crediting a salary of at least S$3,000 through an employer via GIRO, PayNow or FAST."],
             ["Unlock 0.90% p.a. for spending at least S$1,000 in eligible transactions on a Bonus$aver World Mastercard Credit Card or Debit Card."],
             ["Unlock 1.50% p.a. for 6 months when purchasing a new eligible Unit Trust or Online Equity."],
             ["Unlock 2.50% p.a. for 6 months when purchasing a regular premium insurance policy."]])]],

        ["What are the salary crediting requirements for Standard Chartered Bonus$aver bonus interest?",
         [["To qualify for bonus interest on salary credit, there must be a regular inward credit of at least S$3,000 nett of monthly salary into the Bonus$aver Account by the employer through GIRO, PayNow or FAST with Purpose Code “SALA”. The GIRO arrangement must be set up for Singapore Dollar denominated balances, and no interest is received for foreign currency deposited into the account. Salary crediting can be done through instant transfer via PayNow or FAST. The bonus interest rate is 0.90% p.a. on the first S$100,000, credited one month after the salary is credited, and payable monthly on fulfilment of the criteria."]]],

        ["What card spend unlocks bonus interest on a Standard Chartered Bonus$aver Account?",
         [["A minimum card spend of S$1,000, or the equivalent in foreign currency, of eligible card spend per calendar month based on transaction posting date, unlocks a bonus interest rate of 0.90% p.a. on the first S$100,000. Including the prevailing rate, the total interest earned for spend is 0.95% p.a. Bonus interest is credited one month after the card spend transaction month, so for card spending in January bonus interest is paid in February based on January’s Average Daily Balance."]]],

        ["What investment unlocks bonus interest on a Standard Chartered Bonus$aver Account?",
         [["Subscribing to an Eligible Unit Trust with a minimum subscription sum of S$30,000 in a single transaction, or executing a minimum cumulative buy trade volume of S$30,000 in eligible Online Equities via SC Online Trading within a calendar month, unlocks an additional 1.50% p.a. bonus interest for 6 months on the first S$100,000 of the deposit balance. Bonus interest is credited one month after the investment settlement month. Bonus interest for the full 6 months is awarded only if the investments are maintained for the full 6-month holding period; redeeming the Unit Trust or Online Equities during that period forfeits the bonus interest for the remaining months."]]],

        ["What is an Eligible Unit Trust for Standard Chartered Bonus$aver bonus interest?",
         [["An Eligible Unit Trust means a Unit Trust made available through the Bank comprising a minimum investment sum of S$30,000, or its equivalent in another currency, in a single subscription. It does not include exchange traded funds, investments via regular savings plans, or switching transactions."]]],

        ["What are Eligible Online Equities for Standard Chartered Bonus$aver bonus interest?",
         ["Eligible Online Equities refer to shares that are tradable on SC Online Trading, but do not include:",
          B([["Delisted or suspended shares, and shares that are in certificate form."],
             ["Warrants and LPS securities."],
             ["US OTC securities and shares under SGX ETF restrictions."],
             ["Securities sanctioned by the Bank’s internal policies, and shares subject to regulatory and compliance restrictions."]])]],

        ["What insurance purchase unlocks bonus interest on a Standard Chartered Bonus$aver Account?",
         [["Purchasing an Eligible Insurance Policy with a minimum annual premium of S$24,000 unlocks an additional 2.50% p.a. bonus interest on the first S$100,000 of the deposit balance. An Eligible Insurance Policy means a regular premium life insurance policy underwritten by Prudential Assurance Company Singapore (Pte) Ltd and distributed by the Bank, which names the Eligible Customer as the policy owner and has a minimum annual premium of S$24,000. Bonus interest is credited one month after the purchase of the policy and paid for a consecutive period of 6 months."]]],

        ["Who is eligible for a Standard Chartered Bonus$aver Account?",
         [["Applicants must be 18 years and above to be eligible. For joint account applicants, both parties must be 18 years and above. For clients applying for a Bonus$aver account and a Bonus$aver Debit Card the minimum age is 18 years old, and for those applying for a Bonus$aver account and a Bonus$aver Credit Card the age range is 21 to 65 years old."]]],

        ["Can foreigners open a Standard Chartered Bonus$aver Account?",
         [["Yes, foreigners can open a Bonus$aver Account provided they meet the Bank’s requirements. All bonus interest and transactions are in Singapore Dollars. A multi-currency feature allows balances to be held in other currencies, but bonus interest is only earned on SGD balances. Additional funds can also be placed in a foreign currency time deposit account to earn interest."]]],

        ["How is a Standard Chartered Bonus$aver Account opened online?",
         ["A Bonus$aver account can be opened instantly online by applying through MyInfo or Online Banking:",
          B([["New customers: select “Apply now”, pre-fill the application via SingPass (MyInfo), verify the details and submit the application, and the account is opened instantly."],
             ["Existing customers with Online Banking access: log in to Online Banking or SC Mobile, select Apply or Apply for product and search for Bonus$aver, select “Apply now”, verify the details and submit, and the account is opened instantly."]])]],

        ["How much does it cost to open a Standard Chartered Bonus$aver Account, and how many can be held?",
         [["There is zero cost to open a Bonus$aver Account, though other fees apply to the account. A maximum of 3 Bonus$aver Accounts can be held at any one time."]]],

        ["Can a Standard Chartered Bonus$aver Account transact in foreign currency?",
         ["Yes, up to 14 currencies can be transacted in by enabling the multi-currency feature:",
          B([["Step 1: Log on to Standard Chartered Online Banking or SC Mobile."],
             ["Step 2: Click on a currency tile under Bonus$aver."],
             ["Step 3: Choose the currency and complete the transfer."]])]],

        ["How can a joint Standard Chartered Bonus$aver Account be used to maximise interest?",
         [["A joint Bonus$aver account allows clients to combine incomes and transactions to meet the criteria for bonus interest more easily. By crediting both salaries into the joint account, the total credited amount increases, potentially qualifying for higher interest rates. Note that where a joint investment account is used, only the primary account holder of the Bonus$aver account is entitled to bonus interest."]]],

        ["How are foreign currency investments valued for Standard Chartered Bonus$aver bonus interest?",
         [["When investing in eligible Unit Trusts or Online Equities denominated in foreign currencies, the value of such investments is converted into Singapore dollars at the Bank’s prevailing currency exchange rate. To qualify for bonus interest, the converted amount must meet the minimum single subscription amount required. The Bank determines the equivalent Singapore Dollar amount using exchange rates it reasonably considers appropriate."]]],

        ["Does an existing Standard Chartered Bonus$aver investment made before January 2026 still qualify for bonus interest?",
         [["If the investment was purchased before January 2026, the client continues to enjoy the Invest bonus interest of 2.50% p.a. for the remaining bonus interest months, without needing to purchase new eligible Unit Trusts or Online Equities at the higher S$30,000 minimum."]]],

        ["What charges apply to the Standard Chartered Bonus$aver Account?",
         ["The key charges on the Bonus$aver Account are:",
          B([["Fall-below fee: S$5 per month, chargeable if the minimum average daily balance falls below S$3,000 in any particular month."],
             ["Monthly service charge: nil."],
             ["Unarranged overdraft: Prime +5%, minimum S$5."],
             ["Credit card annual fee including GST: S$218, free for the first year."],
             ["Debit card annual fee including GST: S$0."],
             ["Cheque book charges: S$10 for each cheque book requested."],
             ["Early account closure within 6 months: S$30."]])]],

        ["What is the Standard Chartered Bonus$aver sign-up promotion?",
         [["Signing up for a Bonus$aver Account and a Bonus$aver World Mastercard Credit Card gives S$228 cashback when S$50,000 in fresh funds is deposited and maintained upon account opening. The offer is valid from 1 July to 31 August 2026. Terms and conditions apply."]]],
      ]},

      { name: "Cheque and Save Account", url: U.chequesave, qas: [
        ["What are the benefits of the Standard Chartered Cheque and Save Account?",
         ["The Cheque and Save Account offers the following:",
          B([["Free GIRO service, allowing payments by GIRO at no extra charge, and standing instructions for regular payments."],
             ["Cash deposits at Standard Chartered Cash Deposit Machines, and the full range of the Bank’s remittance services."],
             ["Cash withdrawals from Standard Chartered’s extensive ATM network, and NETS payments all over Singapore."],
             ["Account access anytime via SC Mobile or Online Banking, with no lock-in period."],
             ["A consolidated statement for easy account balancing."]])]],

        ["What charges apply to the Standard Chartered Cheque and Save Account?",
         ["The key charges on the Cheque and Save Account are:",
          B([["Minimum initial deposit: S$5,000 for both personal and corporate accounts."],
             ["Fall-below fee: S$7.50 per month if the average daily balance for the month falls below S$5,000."],
             ["Monthly service charge: S$2."],
             ["Unarranged overdraft: Prime +5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30 personal, S$50 corporate."],
             ["Cheque book: S$10 for each cheque book requested."]])]],

        ["Who is eligible to open a Standard Chartered Cheque and Save Account, and what documents are required?",
         ["The Cheque and Save Account is open to Singapore citizens, Singapore permanent residents and foreigners aged 18 years and above. " + DOCS_SGP + " " + DOCS_FOREIGN_INTRO,
          B(DOCS_FOREIGN)]],
      ]},

      { name: "SuperSalary Account", url: U.supersalary, qas: [
        ["What are the benefits of the Standard Chartered SuperSalary Account?",
         ["The SuperSalary Account offers the following:",
          B([["0.01% p.a. interest on balances."],
             ["1% cashback on eligible Mastercard spend with the CashBack debit card, with monthly cashback capped at S$60 per account."],
             ["Money management through Online and Mobile Banking, with SC Mobile available on iPhone, iPad and Android devices."],
             ["Free payment of credit card bills from any bank via Online Banking or Mobile Banking, accepting any Visa or Mastercard cards issued in Singapore."],
             ["Access to over 300 billing organisations via Online Banking, or GIRO arrangements for easy payment of bills."]])]],

        ["What charges apply to the Standard Chartered SuperSalary Account?",
         ["The key charges on the SuperSalary Account are:",
          B([["Minimum initial deposit: S$0, minimum average daily balance: S$0, fall-below fee: S$0, monthly service charge: S$0."],
             ["Non inward credit fee: S$5, charged from the fourth month after account opening if there is zero or less than S$1,500 credit received into the account."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30."],
             ["Debit card annual fee: S$0."],
             ["Cheque book: S$10 for each cheque book requested."]])]],

        ["Who is eligible to open a Standard Chartered SuperSalary Account, and what documents are required?",
         ["The SuperSalary Account is open to Singapore citizens, Singapore permanent residents and foreigners aged 18 years and above. " + DOCS_SGP + " " + DOCS_FOREIGN_INTRO,
          B(DOCS_FOREIGN)]],
      ]},

      { name: "Wealth $aver Account", url: U.wealthsaver, qas: [
        ["What are the benefits of the Standard Chartered Wealth $aver deposit account?",
         ["The Wealth $aver deposit account offers the following benefits:",
          B([["A high interest rate of up to 2.80% p.a. capped at S$2,000,000 for the SGD deposit balance, and up to 4.00% p.a. capped at US$2,000,000 for the USD deposit balance."],
             ["Unlimited 1% cashback on both local and foreign currency Wealth $aver debit card spend, with no minimum spending and no cashback cap."],
             ["S$0 overseas transaction fee on foreign currency spend."],
             ["Access to up to 14 currencies in one single account, converted at attractive FX rates."],
             ["Access to a full suite of wealth expertise, and savings of up to 16% off fuel."]])]],

        ["Who is eligible for a Standard Chartered Wealth $aver deposit account?",
         [["The Wealth $aver deposit account is exclusively for Standard Chartered Priority Banking or Priority Private clients who are at least 18 years old. A client can become a Priority Banking client by opening a Wealth $aver deposit account and depositing a minimum of S$200,000 fresh funds."]]],

        ["What relationship-based SGD interest rates apply to the Standard Chartered Wealth $aver account?",
         ["For Priority Banking clients and above, the SGD interest rate is based on Assets Under Management (AUM):",
          B([["AUM below S$200,000: 0.05% p.a., with no cap on deposit balance."],
             ["AUM from S$200,000 to below S$1,500,000: 0.40% p.a."],
             ["AUM of S$1,500,000 and above: 0.80% p.a."],
             ["The rate applies up to the first S$2,000,000 of balance, and amounts exceeding S$2,000,000 earn 0.05% p.a."]])]],

        ["What relationship-based USD interest rates apply to the Standard Chartered Wealth $aver account?",
         ["The USD interest rate is based on Assets Under Management (AUM):",
          B([["AUM below S$200,000: 0.05% p.a."],
             ["AUM from S$200,000 to below S$1,500,000: 1.00% p.a."],
             ["AUM of S$1,500,000 and above: 2.00% p.a."],
             ["The USD deposit rate is capped at US$2 million of deposit balance."]])]],

        ["What is Assets Under Management for the Standard Chartered Wealth $aver account?",
         ["AUM is defined as the sum of:",
          B([["Total deposits placed with the Bank, including Time Deposits."],
             ["Total value of eligible investment products with the Bank."],
             ["Total premiums paid to date on eligible in-force insurance products purchased through the Bank."]])]],

        ["What is the Standard Chartered Wealth $aver Wealth Booster Promotion?",
         ["The Wealth Booster Promotion unlocks bonus interest for 6 months when, within a calendar month during the Promotion Period, the client either purchases Eligible Insurance Policies with a Total Annual Premium of at least S$24,000, or subscribes and settles trades for Eligible Unit Trusts, Bonds and Structured Notes with a Total Subscription Amount of at least S$200,000. The rates are:",
          B([["New-to-deposit client: 1.00% p.a. for 6 months."],
             ["Existing-to-deposit client: 0.50% p.a. for 6 months."],
             ["The 6 months start from the month of inception of the Eligible Insurance Policies or the investment subscription month, and Promotional Bonus Interest is credited 1 month after the calendar month eligible for it."]])]],

        ["Who is a “New-to-deposit” client for the Standard Chartered Wealth $aver Wealth Booster Promotion?",
         [["A New-to-deposit client is one applying for the Wealth $aver account for the first time during the Promotion Period who does not hold an existing Standard Chartered current, cheque or savings account at the date of Wealth $aver account opening. The client must also complete their Investment or Insurance purchase by the Wealth $aver account opening month plus the following one calendar month. For example, if the Wealth $aver account is opened in August 2026, the client must make a new purchase of Insurance or Investment by 30 September 2026."]]],

        ["What are Eligible Investment Products under the Standard Chartered Wealth Booster Promotion?",
         ["The following are Eligible Investment Products:",
          B([["Eligible Unit Trust: a unit trust distributed by the Bank, excluding exchange traded funds and switching transactions."],
             ["Eligible Bonds: successful subscription of bonds issued and/or distributed by the Bank."],
             ["Eligible Structured Notes: successful subscription of structured notes issued and/or distributed by the Bank, excluding Premium Currency Investments and Commodity Linked Structured Investment."],
             ["The aggregate subscription for all Eligible Investment Products purchased through the Bank within a calendar month during the Promotion Period must be at least S$200,000 or its equivalent in another currency. For Non-Accredited Investors, eligible Bonds and Structured Notes are subject to minimum investment amounts."]])]],

        ["What are Eligible Insurance Policies under the Standard Chartered Wealth Booster Promotion?",
         [["Eligible Insurance Policies are regular premium or single premium life insurance policies underwritten by Prudential Assurance Company Singapore (Pte) Limited, with the exception of PRUShield and single premium life insurance policies purchased using SRS funds, which name the Wealth $aver accountholder as the policy holder. The aggregated annualised premiums payable for all Eligible Insurance Policies purchased through the Bank within a calendar month during the Promotion Period must be at least S$24,000 or its equivalent in another currency."]]],

        ["Must a Standard Chartered Wealth $aver account be opened before investing to qualify for the Wealth Booster Promotion?",
         [["Yes. A Wealth $aver deposit account must be opened before the subscription of Eligible Investment Products or the purchase of Eligible Insurance Policies in order to qualify for the Promotional Bonus Interest."]]],

        ["What is the Standard Chartered Wealth $aver Top-Up Promotion?",
         ["The Top-Up Promotion unlocks bonus interest for 3 months and requires registration. The rates are:",
          B([["SGD: top-up fresh funds of S$1M earns 1.00% p.a., and top-up fresh funds of S$200K earns 0.50% p.a."],
             ["USD: top-up fresh funds of US$1M earns 1.00% p.a., and top-up fresh funds of US$200K earns 0.50% p.a."]])]],

        ["When and how is interest calculated on a Standard Chartered Wealth $aver account?",
         [["Interest is credited to the Wealth $aver deposit account monthly by the last day of the following month, or at other regular intervals determined by Standard Chartered. For Priority Banking or Priority Private clients at the point of interest computation, the average Assets Under Management for the month determines the interest rate assigned to the account each month. Interest is earned on the average daily balance for Singapore Dollar denominated balances. A client who is not a Priority Banking or Priority Private client earns only 0.05% p.a., and it may take up to 7 working days to process a Priority Banking or Priority Private sign-up request after the eligibility criteria are fulfilled."]]],

        ["What charges apply to the Standard Chartered Wealth $aver account?",
         ["The key charges on the Wealth $aver account are:",
          B([["Fall-below fee: S$5 per month."],
             ["Unarranged overdraft: Prime +5%, minimum S$5."],
             ["Cheque book charges: S$10."],
             ["Early account closure within 6 months: S$30."],
             ["Except for unarranged overdraft charges, all other fees are waived for as long as the client remains a Priority Banking client."]])]],

        ["What documents are required to open a Standard Chartered Wealth $aver account?",
         ["Applicants must be 18 years old and above and Priority Banking or above. The documents required are:",
          B([["SingPass holders applying with MyInfo: no documents are required."],
             ["Singapore citizens and permanent residents applying at a branch: original Singapore NRIC."],
             ["Foreigners applying at a branch: a copy of a passport with at least 6 months’ validity, a copy of the Employment Pass, and any one of the latest utility, rates or tax bill; latest bank or credit card statement including e-Statements; rental agreement showing the address; latest mobile phone or pay-TV statement; or a letter from the employer stating the current address."]])]],
      ]},
    ],
  },

  // ========================================================== TIME DEPOSITS
  {
    name: "Time Deposits",
    subs: [
      { name: "Singapore Dollar Time Deposit", url: U.tdSgd, qas: [
        ["What is the Standard Chartered Singapore Dollar Time Deposit?",
         ["The Singapore Dollar Time Deposit is a fixed deposit offering:",
          B([["Competitive interest rates."],
             ["Flexible tenors ranging from 1 to 24 months."],
             ["Instant application via SC Mobile or Online Banking."]])]],

        ["What is the Standard Chartered Singapore Dollar Time Deposit Fresh Funds promotion?",
         [["From 4 August 2026 to 11 August 2026, a promotional rate applies on a Singapore Dollar Time Deposit with a minimum of S$25,000 in Fresh Funds, for a 6-month tenor. The rates are 1.30% p.a. for Personal Banking, 1.40% p.a. for Priority Banking, and 1.60% p.a. for Priority Private. To qualify, Fresh Funds must be transferred to a Standard Chartered deposit account and the time deposit placed by logging in to Online Banking or SC Mobile. The promotional interest rates are only applicable if the Time Deposit is held until the maturity of the tenure."]]],

        ["What are “Fresh Funds” for the Standard Chartered Singapore Dollar Time Deposit promotion?",
         [["Fresh funds refer to funds not originating from any existing account with Standard Chartered Bank (Singapore) Limited, and funds that are not withdrawn and re-deposited within the last 30 days of opening the Time Deposit."]]],

        ["How is a Standard Chartered Singapore Dollar Time Deposit placed via SC Mobile?",
         ["Fresh funds should first be transferred to an existing savings or current account, then:",
          B([["Step 1: Log in to SC Mobile."],
             ["Step 2: Click on “Discover”, located at the bottom navigation bar."],
             ["Step 3: Select “Deposits” then “Time Deposit”."],
             ["Step 4: Select “Singapore Dollar Time Deposit”."],
             ["Step 5: Complete the application. Details of the Time Deposit placement can then be viewed under Account Summary."]])]],

        ["How is a Standard Chartered Singapore Dollar Time Deposit placed via Online Banking?",
         ["Fresh funds should first be transferred to an existing savings or current account, then:",
          B([["Step 1: Log in to iBanking and select “Apply” on the top navigation bar."],
             ["Step 2: Follow the steps to authenticate the login."],
             ["Step 3: Select the preferred “Time Deposit” account."],
             ["Step 4: Fill in the details to complete and submit the application. Details of the Time Deposit placement can then be viewed under Account Summary."]])]],

        ["Will statements be received for a Standard Chartered Singapore Dollar Time Deposit placement?",
         [["A Time Deposit advice is generated and sent upon placement. Clients currently receiving a Consolidated Statement may request to link the Time Deposit account to their Consolidated Statement."]]],

        ["What happens if a Standard Chartered Singapore Dollar Time Deposit matures on a Sunday or public holiday?",
         [["If the maturity date falls on a non-working day, such as a Sunday or a public holiday, the maturity date is moved to the next working day and interest is paid accordingly for the non-working day."]]],

        ["Can a Standard Chartered Singapore Dollar Time Deposit be withdrawn before maturity?",
         [["Yes, but if any withdrawals are made from the Time Deposit before the maturity date, any interest payable on the Time Deposit is calculated at Standard Chartered’s sole discretion. In some circumstances, no interest at all may be received on the Time Deposit."]]],

        ["What phishing scam has Standard Chartered warned about for Fixed Deposit rates?",
         [["Standard Chartered has warned clients to be on high alert for a targeted SMS phishing scam from a +65 number offering promotional Fixed Deposit rates. The Bank would never contact clients using a +65 number or direct them to a personal mobile number for more information. Personal particulars should never be disclosed to anyone."]]],
      ]},

      { name: "Foreign Currency Time Deposit", url: U.tdFcy, qas: [
        ["What is the Standard Chartered Foreign Currency Time Deposit?",
         ["The Foreign Currency Time Deposit is a fixed deposit offering:",
          B([["Nine foreign currencies: USD, GBP, AUD, NZD, EUR, CAD, HKD, CHF and CNH."],
             ["Flexible tenors from 1 to 24 months, locking in a fixed interest rate."],
             ["Instant application via SC Mobile or Online Banking."]])]],

        ["What is the Standard Chartered USD Time Deposit Fresh Funds promotion?",
         [["From 4 August 2026 to 11 August 2026, a minimum fresh funds deposit of USD25,000 on a 9-month tenor earns a promotional interest rate of 3.90% p.a. for Personal Banking, 4.00% p.a. for Priority Banking, and 4.20% p.a. for Priority Private Banking."]]],

        ["What are the minimum placement amounts for a Standard Chartered Foreign Currency Time Deposit?",
         ["The minimum placement amounts by currency are:",
          B([["USD 5,000, GBP 5,000, AUD 5,000, EUR 5,000 and NZD 5,000."],
             ["CAD 25,000, HKD 25,000, CHF 25,000 and CNH (Renminbi Offshore) 25,000."]])]],

        ["How is a Standard Chartered Foreign Currency Time Deposit placed via SC Mobile?",
         ["Fresh funds should first be transferred to an existing FCY account, then:",
          B([["Step 1: Log in to SC Mobile."],
             ["Step 2: Click on “Discover”, located at the bottom navigation bar."],
             ["Step 3: Select “Deposits” then “Time Deposit”."],
             ["Step 4: Select “Foreign Currency Time Deposits”."],
             ["Step 5: Complete the application. Details of the Time Deposit placement can then be viewed under Account Summary."]])]],

        ["Which accounts can fund a Standard Chartered Foreign Currency Time Deposit?",
         [["Fresh funds must be transferred to an existing foreign currency account before applying for a Foreign Currency Time Deposit. Clients without an eligible FCY account can apply for a Bonus$aver or Wealth $aver account with multi-currency features, or a USD$aver account."]]],

        ["Can a Standard Chartered Foreign Currency Time Deposit be withdrawn before maturity?",
         [["Yes, but if any withdrawals are made from the Time Deposit before the maturity date, any interest payable is calculated at Standard Chartered’s sole discretion, and in some circumstances no interest at all may be received. Penalty charges and replacement cost apply for all premature upliftment. Applicable charges are published in the ", ["Standard Chartered Singapore Pricing Guide", L.pricingGuide], "."]]],

        ["Will statements be received for a Standard Chartered Foreign Currency Time Deposit placement?",
         [["A Time Deposit advice is generated and sent upon placement. Clients currently receiving a Consolidated Statement may request to link the Time Deposit account to their Consolidated Statement."]]],

        ["What happens if a Standard Chartered Foreign Currency Time Deposit matures on a Sunday or public holiday?",
         [["If the maturity date falls on a non-working day, such as a Sunday or a public holiday, the maturity date is moved to the next working day and interest is paid accordingly for the non-working day. The public holiday calendar follows the respective country of the Time Deposit currency."]]],
      ]},

      { name: "Sustainable Time Deposit", url: U.tdSustainable, qas: [
        ["What is a Standard Chartered Sustainable Time Deposit?",
         [["Sustainable Time Deposits are fixed deposits which allow clients to have their capital referenced against sustainable loans and projects of Standard Chartered. These assets include green financing, sustainable infrastructure projects, microfinance and business banking. The deposits earn high interest throughout the deposit tenure."]]],

        ["What is the Standard Chartered Sustainable Time Deposit promotion?",
         [["From 4 August 2026 to 11 August 2026, a promotional rate applies on a Sustainable Time Deposit with a minimum placement of S$25,000 in fresh funds, for a 6-month tenure. The rates are 1.30% p.a. promotional, 1.40% p.a. Priority Banking preferential, and 1.60% p.a. Priority Private Banking preferential. To qualify, fresh funds must be transferred to a Standard Chartered deposit account and the time deposit placed by logging in to Online Banking or SC Mobile."]]],

        ["What assets fit Standard Chartered’s Green and Sustainable Product Framework?",
         ["The themes within the Green and Sustainable Product Framework include COVID healthcare, sanitation and food security, along with others such as renewable energy, energy efficiency, access to water and water management, sustainable infrastructure, climate change adaptation, access to finance, healthcare, education and partnerships. Examples are:",
          B([["COVID healthcare: financing to equip, operate and add capacity and efficiency to essential healthcare facilities."],
             ["Renewable energy: financing the generation of electricity from wind, solar, hydropower, waste to energy and geothermal."],
             ["Education: construction of public schools and universities, construction of student housing, and training for educational professionals."]])]],

        ["How does Standard Chartered ensure Sustainable Time Deposit monies go to the right assets?",
         [["Standard Chartered has mapped its existing business against the Sustainable Development Goals to create a Green and Sustainable Product Framework detailing what the organisation views as sustainable activities that it finances. The framework has received industry accreditation from Sustainalytics and is reviewed independently every year to ensure the latest developments and trends are incorporated. The framework sets the basis for future sustainable products to help fund the SDGs across the Bank’s footprint markets, especially in developing markets."]]],

        ["What is the scale of Standard Chartered’s Sustainable Finance Portfolio?",
         ["Standard Chartered reports the following on its Sustainable Finance Portfolio, which holds USD 12.9bn in sustainable assets, up 40% from 2021:",
          B([["Social: over 700,000 microfinance loans enabled."],
             ["Environmental: 1.59 million tonnes of CO2 saved in the last year from operational assets and those in construction."],
             ["Financial: nearly 20,000 SME loans disbursed."],
             ["Geographical: over 90% of Sustainable Finance assets are located in Asia, Africa and the Middle East."]])]],

        ["How are Standard Chartered Sustainable Time Deposits referenced against assets?",
         [["Sustainable Time Deposits are referenced against assets held in aggregate by the SCB Group, whether existing at the dates when the deposits are placed or in the future, that SCB Group deems sustainable in accordance with its externally verified Green and Sustainable Product Framework. The assets referenced against the Sustainable Time Deposits are at least equal to or greater in value than the Sustainable Time Deposits. SCB Group means Standard Chartered PLC and its subsidiaries and affiliates. The placement of Sustainable Time Deposits is subject to acceptance, including any limits on placement amounts, by the Bank."]]],
      ]},
    ],
  },

  // ============================================== ACCOUNT FEATURES AND GUIDANCE
  {
    name: "Account Features and Guidance",
    subs: [
      { name: "Choosing and Opening an Account", url: U.hub, qas: [
        ["Which Standard Chartered savings account gives high interest?",
         [["The Bonus$aver Account is a savings account which offers high interest rates, where bonus interest can be earned from the very first dollar."]]],

        ["How is a bank account opened with Standard Chartered Singapore?",
         ["An account is opened in 3 simple steps:",
          B([["Step 1: Validate eligibility. Check the eligibility requirements for opening an account. For selected savings accounts, the account can only be opened at 16 years old and above. Foreigners are eligible as well. Check the respective product pages for details."],
             ["Step 2: Submit documents. Different account applications require different supporting documents based on citizenship status, covering Singapore Citizens, Permanent Residents and Foreigners. Singaporeans and Permanent Residents need only their SingPass to apply via MyInfo, with no additional documents required."],
             ["Step 3: Start banking. Download SC Mobile to bank from anywhere, anytime."]])]],

        ["Can a Standard Chartered fixed deposit account be opened online?",
         ["Yes, a fixed deposit account can be opened instantly through Online Banking:",
          B([["Step 1: Log in to Online Banking."],
             ["Step 2: Click “Apply” on the top navigation bar."],
             ["Step 3: Select “Time Deposit Account”."],
             ["Step 4: Fill up the application form and select the tenure of choice."],
             ["Once the Time Deposit request is submitted successfully, the new placement is reflected under the Account Summary."]])]],

        ["Is there a Standard Chartered savings account for young adults?",
         [["JumpStart is a savings account specially curated for young adults between 18 and 26 years old. The key features of the account include up to 1.5% p.a. interest on the first $50,000, with no fees, minimum deposit, lock-in period or salary crediting requirement."]]],

        ["Are Standard Chartered Singapore deposits insured?",
         [["Yes. Generally, all Singapore dollar deposits in savings, fixed deposit and current accounts are covered under the Deposit Insurance Scheme. With effect from 1 April 2024, Singapore dollar deposits held by non-bank depositors in eligible savings, fixed deposit and current accounts are insured up to S$100,000 per depositor in accordance with the Deposit Insurance and Policy Owners’ Protection Schemes Act (CAP77B). Foreign currency deposits, dual currency investments, structured deposits and other investment products are not insured."]]],
      ]},

      { name: "Multi-Currency Feature", url: U.mca, qas: [
        ["What is the Standard Chartered Multi-Currency Feature?",
         [["The Multi-Currency Feature is applicable to the Bonus$aver and Wealth $aver Accounts. It allows the account holder to transact in up to 14 currencies from their account, with zero overseas transaction fees, and to convert at competitive FX rates with LiveFX. Overseas transaction fees here refer to fees charged by Standard Chartered Bank (Singapore) Limited only; other banks’ fees may still apply."]]],

        ["Who can enable the Standard Chartered Multi-Currency Feature?",
         ["Eligibility for the Multi-Currency Feature is:",
          B([["Account holders of a personal Bonus$aver or Wealth $aver account."],
             ["The account holder must be registered with SC Mobile to view the option to enable this feature."],
             ["It is available for new or existing personal Bonus$aver or Wealth $aver account holders."]])]],

        ["How is the Standard Chartered Multi-Currency Feature enabled?",
         ["An existing Bonus$aver or Wealth $aver account holder enables the feature as follows:",
          B([["Step 1: Log on to Standard Chartered Online Banking or the Mobile Banking app."],
             ["Step 2: Tap “Activate Multi-currency” under the Bonus$aver or Wealth $aver Account and select the chosen currency."],
             ["Step 3: Transfer funds from another Standard Chartered account to the chosen currency."],
             ["Step 4: Complete the transfer, and the chosen currency is enabled for use."]])]],

        ["Who benefits from the Standard Chartered Multi-Currency Feature?",
         ["The Multi-Currency Feature suits the following:",
          B([["Online shoppers, who can pay like a local at checkout using a multi-currency enabled debit card with zero overseas transaction fees."],
             ["Expatriates, who can send money to loved ones with no incremental fees."],
             ["Investors, who can access funds in 14 currencies with one account and take advantage of competitive FX rates."],
             ["Overseas students, who can transact like a local and spend in foreign currency with zero overseas transaction fees."]])]],

        ["Which currencies cannot be sent by Outward Telegraphic Transfer on SC Mobile or Online Banking?",
         [["Outward Telegraphic Transfer via SC Mobile or Online Banking is not available for DKK, SEK and ZAR. Transfers in these currencies must be arranged at a Standard Chartered branch."]]],
      ]},

      { name: "Getting Your First Credit Card", url: U.jumpstartStory, qas: [
        ["What should be considered when choosing a first credit card?",
         ["Standard Chartered advises spending time researching different credit card options and their benefits before deciding, looking at three factors:",
          B([["Rewards and incentives: cards can offer perks such as 1.5% cashback on eligible spending with no cap and no minimum spend, airline miles, or points redeemable for cash, discounts or merchandise."],
             ["Interest rates: the accrued amount payable on a credit card’s unpaid balance. In Singapore, credit card interest rates can go up to about 26.9%, so debt can snowball significantly if payments are not made quickly."],
             ["Annual rates and fees: cards may come with annual fees, card replacement fees and withdrawal fees. Some cards do not charge annual fees, and some do not charge until the second year of membership."]])]],

        ["Why is paying only the minimum sum on a credit card a problem?",
         [["When a credit card bill is received, there is an option to pay only a minimum amount, usually 3% of the outstanding amount or S$50, whichever is higher. Because of compounding interest, paying only the minimum can cost significantly more over time. The advice is to always pay back in full, or at the very least more than the minimum sum."]]],

        ["How can automatic payments help manage a credit card?",
         [["Setting up automatic payments ensures bills are always paid on time, as the amount is directly debited from a bank account, provided no big-ticket items are purchased that cannot be paid for in the short term. This in turn means no interest is payable."]]],

        ["Why should cash advances be avoided on a credit card?",
         [["Cash advances come with high fees and interest rates, often even higher than regular credit card interest rates. If the cash advance cannot be repaid, it further adds to the credit card balance and compounding interest. For emergency funds, alternatives such as getting deferments on bills or a personal loan, where interest fees are typically lower, may be worth considering."]]],

        ["Why should credit card monthly statements be monitored?",
         [["A credit card statement is received every month and is a good way to track spending and confirm the budget was kept to. It is good practice to check monthly statements for any payment discrepancies, especially if someone else has held the credit card or has access to the card details saved for mobile purchases. Contactless payments are sometimes assumed to cause identity theft, but with the layers of security in place this is almost impossible."]]],

        ["Why should a credit score be checked at least once a year?",
         [["A credit score is an indicator of how likely a person is to repay their debts, and lenders use it to assess whether they are a good risk. A good credit score can secure lower rates on future loans and improve the chance of approval for other cards. A credit file copy can be requested online, at any Singapore Post branch, or from the Credit Bureau Singapore. The general ground rules are to pay bills on time and not max out credit limits."]]],

        ["What is the risk of using a 0% instalment plan on a credit card?",
         [["Most credit cards offer a 0% instalment plan, which means that for an item costing S$1,200 the cardholder pays S$100 per month for 12 months. However, if the monthly fees cannot be paid, interest still accrues on the credit card balance. Purchases should only be made if they can be paid for within a set timeframe, after factoring in salary and monthly expenditure."]]],
      ]},
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Save FAQ",
  docTitle: "SCB Bank_Save FAQ",
  description: "Standard Chartered Singapore Save FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
console.log("categories:", CATEGORIES.length, "subcategories:", CATEGORIES.reduce((n, c) => n + c.subs.length, 0));
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Save_FAQ.docx", buf);
  console.log("written", buf.length);
});
