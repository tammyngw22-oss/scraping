const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

const BR = "https://www.sc.com/sg/borrow/";
const U = {
  hub: BR,
  cashone: BR + "loans/cashone/",
  ccft: BR + "loans/balance-transfer/",
  dcp: BR + "loans/debt-consolidation/",
  mortgages: BR + "mortgages/",
  mortgageone: BR + "mortgages/mortgageone/",
  repricing: BR + "mortgages/loanrepricing/",
  sora: BR + "mortgages/sora/",
  green: BR + "mortgages/green-mortgage/",
};

const L = {
  pricingGuide: "https://www.sc.com/sg/pricing-guide/",
  mortgageTerms: "https://av.sc.com/sg/content/docs/scb-mortgage-facility-tc.pdf",
  greenMarkDir: "https://www.sleb.sg/Building/GreenMarkBuildingsDirectory",
  bonussaver: "https://www.sc.com/sg/save/current-accounts/bonussaver/",
  simplyCash: "https://www.sc.com/sg/credit-cards/simply-cash-credit-card/",
  easypay: "https://www.sc.com/sg/bank-with-us/manage-your-finances/easypay/",
  easybill: "https://www.sc.com/sg/bank-with-us/manage-your-finances/easybill/",
  sustainableTd: "https://www.sc.com/sg/save/time-deposits/sustainable-time-deposit/",
};

const CATEGORIES = [
  // ==================================================== PERSONAL FINANCING
  {
    name: "Personal Financing",
    subs: [
      { name: "Lending Solutions Overview", url: U.hub, qas: [
        ["Which personal financing solutions does Standard Chartered Singapore offer?",
         ["Standard Chartered Singapore offers three personal financing solutions:",
          B([["Credit Card Funds Transfer: instant cash at 0% interest over a 3 to 12 month tenure, with a processing fee from 0.90% (EIR from 2.06% p.a.), with instant approval and disbursement to a bank account or another bank’s credit card."],
             ["CashOne Personal Loan: long-term flexible financing over a 1 to 5 year tenure, with low interest rates from 0.90% p.a. (EIR from 1.75% p.a.), with instant approval and disbursement to a bank account of choice."],
             ["Debt Consolidation Plan: consolidates outstanding credit card bills and personal loans across multiple banks into a single loan, over a 3 to 10 year tenure, with a credit card with annual fee waived provided for daily use."]])]],

        ["Which mortgage products does Standard Chartered Singapore offer?",
         ["Standard Chartered Singapore offers the following mortgage products:",
          B([["MortgageOne, which offsets loan interest from interest earned on deposits."],
             ["Loan Repricing, which gives lower interest rates on an existing mortgage loan."],
             ["Green Mortgage, which supports a positive impact on society and the environment."]])]],

        ["What can a Standard Chartered CashOne Personal Loan be used for?",
         ["Standard Chartered lists the following uses for a CashOne Personal Loan:",
          B([["Home renovation, to fund a renovation upfront."],
             ["Parenthood, to cover the costs of a growing family."],
             ["A dream vacation, to book a trip without draining savings."],
             ["Money management, to consolidate debt into one repayment."],
             ["A wedding, to pay for the big day."]])]],
      ]},

      { name: "CashOne Personal Loan", url: U.cashone, qas: [
        ["What is the Standard Chartered CashOne Personal Loan?",
         ["CashOne is a personal loan with the following features:",
          B([["Low rates from 0.90% p.a. (EIR from 1.75% p.a.)."],
             ["Instant approval and disbursement, subject to eligibility criteria."],
             ["No late payment charges, an exclusive benefit for customers who stay on track with payments."],
             ["A flexible repayment tenure from 1 to 5 years."]])]],

        ["How is the Standard Chartered CashOne Effective Interest Rate calculated?",
         [["The EIR of from 1.75% p.a. is based on a 5-year tenure and does not yet include the first-year annual fee of S$199 or any cashback received. Taking the first-year annual fee of S$199 into account, the EIR is 2.15% p.a. for an average loan amount of S$20,000 at a 5-year tenure. The applied interest rate of 0.90% p.a. is the lowest rate, and the rate offered in an application is based on the applicant’s credit profile as determined by the Bank."]]],

        ["Why does the Applied Rate differ from the Effective Interest Rate on a Standard Chartered CashOne Personal Loan?",
         [["The Applied Rate assumes the loan principal stays constant across the tenure. The Effective Interest Rate reflects the true cost of borrowing as the principal declines over time, and also factors in the first-year annual fee of S$199, so it is always higher than the Applied Rate."]]],

        ["How is total interest calculated on a Standard Chartered CashOne Personal Loan?",
         [["Interest is calculated using the front-end add-on method: principal loan amount multiplied by the approved applied rate multiplied by the full tenure. While the monthly instalment stays constant, the interest portion is highest at the start of the loan and decreases gradually over the tenure."]]],

        ["What are the Standard Chartered CashOne SG61 cashback tiers for new credit card clients?",
         ["For clients new to Standard Chartered Credit Cards, the cashback tiers are:",
          B([["S$18,000 to S$29,999: 0.61% cashback, up to S$183."],
             ["S$30,000 to S$49,999: 1.61% cashback, up to S$805."],
             ["S$50,000 to S$209,999: 2.61% cashback, up to S$5,481."],
             ["S$210,000 and above: 3.61% cashback, up to S$9,025."]])]],

        ["What are the Standard Chartered CashOne SG61 cashback tiers for existing credit card clients?",
         ["For clients existing to Standard Chartered Credit Cards, the cashback tiers are:",
          B([["S$18,000 to S$29,999: 0.61% cashback, up to S$183."],
             ["S$30,000 to S$49,999: 1.61% cashback, up to S$805."],
             ["S$50,000 and above: 2.61% cashback, up to S$6,525."]])]],

        ["What conditions apply to Standard Chartered CashOne cashback?",
         [["Cashback is rounded down to 2 decimal places and is for approved loan tenures of 3 to 5 years only. The applicant must have a Standard Chartered Savings or Current Account in SGD currency for cashback crediting. The CashOne SG61 promotions run from 13 July 2026 to 30 September 2026."]]],

        ["Who is eligible for a Standard Chartered CashOne Personal Loan?",
         ["Eligibility for CashOne is:",
          B([["Singaporeans and Permanent Residents: age 21 years and above, with a minimum annual income of S$30,000."],
             ["Foreigners: age 21 years and above, with a minimum annual income of S$90,000, and must hold a Singapore Employment Pass."],
             ["Existing Standard Chartered credit cardholders need no documents unless their income or employment has changed."]])]],

        ["What documents do Singaporeans and Permanent Residents need for a Standard Chartered CashOne Personal Loan?",
         ["No documents are required for SingPass holders applying via MyInfo. Otherwise:",
          B([["Salaried or partial commission-based: copy of NRIC front and back, plus either the latest computerised payslip or the latest 6 months’ CPF Contribution History Statement."],
             ["100% commission-based: copy of NRIC front and back, plus either the latest Income Tax Notice of Assessment or the latest 3 months’ Commission Statement from the same employer."],
             ["Self-employed: copy of NRIC front and back, and the latest Income Tax Notice of Assessment."]])]],

        ["What documents do foreigners need for a Standard Chartered CashOne Personal Loan?",
         ["Foreign applicants must provide a copy of a passport with at least 6 months’ validity, a copy of their Employment Pass, and one proof of address, which may be a latest utility, rates, bank, credit card, or mobile or pay-TV statement, a rental agreement, an employer letter, or a government-issued document such as from IRAS, CPF or ICA. In addition:",
          B([["Salaried: latest computerised payslip. The latest Income Tax Notice of Assessment should also be submitted to be considered for a higher loan amount."],
             ["100% commission-based: latest Income Tax Notice of Assessment, or latest 3 months’ Commission Statement from the same employer."],
             ["Self-employed: latest Income Tax Notice of Assessment."]])]],

        ["What is the annual fee on a Standard Chartered CashOne Personal Loan?",
         [["A first-year annual fee of S$199 is deducted from the approved loan amount. From the second year onwards, until the expiry of the instalment tenure, S$50 is charged annually, and only if instalments were not paid in full by their due dates for the preceding 12 consecutive calendar months."]]],

        ["What fees and charges apply to a Standard Chartered CashOne Personal Loan?",
         ["The fees and charges on CashOne are:",
          B([["Early redemption fee: S$150 or 3% of the outstanding principal, whichever is higher."],
             ["Change of tenure, per change and subject to approval: S$50."],
             ["Late payment charges: S$100."],
             ["Finance charges, calculated daily: EIR 29.9% p.a. (0.082% per day)."],
             ["Default interest, which applies if 2 minimum payments are late within 6 consecutive months: EIR plus 4% p.a."]])]],

        ["How quickly are Standard Chartered CashOne Personal Loan funds disbursed?",
         [["Applying via MyInfo can mean instant approval. Funds disbursed to an existing Standard Chartered Current or Cheque & Save Account are received instantly upon approval. Funds disbursed to a non-Standard Chartered account are typically received within 15 minutes, subject to FAST participation and approval cut-off times."]]],

        ["Does a Standard Chartered CashOne Personal Loan affect an existing credit card limit?",
         [["Yes, for an existing Standard Chartered principal credit cardholder. CashOne is drawn down as a Credit Card Instalment Loan against the existing credit limit, which frees up again as the instalment loan is repaid."]]],

        ["Does taking a Standard Chartered personal loan affect a credit score?",
         [["Yes. Timely payments can improve a credit score, while missed payments or a high debt load could negatively impact it."]]],

        ["Why might an approved Standard Chartered CashOne loan amount be lower than requested?",
         [["The approved amount is subject to the Bank’s review. Existing Standard Chartered credit cardholders are capped at 98% of their available credit card limit. New customers may be issued a Platinum Visa Credit Card if their annual income is above S$30,000. Following the review, the approved loan amount, which may differ from the requested amount, is directly disbursed to the designated account stated in the application."]]],

        ["Why does the actual Standard Chartered CashOne interest rate differ from the calculator rate?",
         [["The rate shown in an online application is based on the applicant’s individual credit profile, so it may differ from the illustrative rate shown in the calculator and from rates shown to other borrowers."]]],

        ["What are the late repayment penalties on a Standard Chartered CashOne Personal Loan?",
         [["A flat late payment fee of S$100 applies for every late month, plus additional interest on the full outstanding balance calculated daily at the minimum EIR of 29.9% p.a., which is 0.082% per day, for as long as the payment remains overdue."]]],

        ["How can a Standard Chartered CashOne borrower qualify for minimum monthly repayments without late penalties?",
         [["A borrower can switch to minimum payments after paying their full monthly instalment consecutively for 6 months. The minimum payment due is the greater of S$50 or 1% of the approved monthly principal instalment amount, plus interest, fees and charges, and any overlimit or past-due amounts. Outstanding balances below S$50 must be repaid in full."]]],

        ["Why might a Standard Chartered CashOne loan tenure be extended mid-way through the plan?",
         [["This happens if a minimum payment due is late or missed twice within a consecutive 6-month period. An additional 4% p.a. is added to the EIR from the next statement date, and the tenure is extended to keep the monthly instalment consistent. The additional interest is removed once minimum payments are made on time for 6 consecutive months."]]],

        ["What repayment methods are available for a Standard Chartered CashOne Personal Loan?",
         [["Repayments can be made through Online Banking, the SC Mobile app, AXS and ATM machines islandwide, any Standard Chartered branch, or GIRO for automatic monthly repayment."]]],

        ["Where is a Standard Chartered CashOne outstanding loan amount found?",
         [["The outstanding loan amount is shown on the monthly statement, or via the “Instalments Loan Summary” under eStatements in Online Banking or the SC Mobile app."]]],

        ["How is a Standard Chartered CashOne Personal Loan settled in full early?",
         [["A closure request is submitted via Online Banking or the SC Mobile app, under Help & Services then Account Management. An early redemption fee of S$150 or 3% of the outstanding principal, whichever is higher, applies."]]],
      ]},

      { name: "Credit Card Funds Transfer", url: U.ccft, qas: [
        ["What is a Standard Chartered Credit Card Funds Transfer?",
         [["A Credit Card Funds Transfer lets a client cash out their available credit limit from their existing Standard Chartered Credit Card at 0% interest for their preferred tenure, with a one-time low processing fee instead of ongoing interest charges. New clients without a Standard Chartered card receive a Simply Cash Credit Card, with an approved Credit Card Funds Transfer of up to 90% of the approved credit limit, plus 1.5% cashback on everyday spend."]]],

        ["What tenures are available on a Standard Chartered Credit Card Funds Transfer?",
         [["There are 4 tenures at 0% interest p.a.: 3 months, 6 months, 9 months and 12 months. The repayment tenure cannot be changed once the transfer is approved."]]],

        ["What are the minimum and maximum Standard Chartered Credit Card Funds Transfer amounts?",
         [["The minimum amount is S$1,000 and the maximum is 90% of the total assigned credit limit. Borrowing is possible up to S$250,000, subject to the available credit limit. For example, a S$10,000 credit limit caps the transfer at S$9,000."]]],

        ["What is the Standard Chartered Credit Card Funds Transfer SG61 cashback offer?",
         ["The SG61 promotion, valid 13 July to 30 September 2026, gives up to 1.61% uncapped cashback on an approved amount of S$20,000 and above on a 12-month tenure:",
          B([["New to Standard Chartered Credit Card: 1.61% cashback."],
             ["Existing to Standard Chartered Credit Card and new to Credit Card Funds Transfer: S$161."],
             ["Existing to Standard Chartered Credit Card and existing to Credit Card Funds Transfer: S$61."],
             ["A Standard Chartered Savings or Current Account in SGD currency is required for cashback crediting."]])]],

        ["Who is eligible for a Standard Chartered Credit Card Funds Transfer?",
         ["Eligibility is as follows:",
          B([["Singaporeans and Permanent Residents: aged 21 and above, with a minimum annual income of S$30,000."],
             ["Foreigners: aged 21 and above, with a minimum annual income of S$90,000, and must hold a valid Singapore Employment Pass."],
             ["Existing Standard Chartered cardholders need no documents unless their income has changed."],
             ["Not eligible: supplementary cardholders, corporate cardholders, and cardholders with a S$500 credit limit."]])]],

        ["What documents are required for a Standard Chartered Credit Card Funds Transfer?",
         ["Documents required depend on the applicant:",
          B([["New to Standard Chartered via MyInfo: details are retrieved instantly via SingPass with no documents to upload."],
             ["Salaried: NRIC and latest payslip, or latest 6-month CPF contribution statement. A Notice of Assessment is required for higher transfer amounts."],
             ["100% commission-based: NRIC and latest Notice of Assessment, or latest 3-month commission statement."],
             ["Self-employed: NRIC and latest Notice of Assessment."],
             ["Foreign applicants must additionally provide a passport valid for 6 months or more, a copy of the Employment Pass, and one proof of address."]])]],

        ["Which banks can a Standard Chartered Credit Card Funds Transfer draw from?",
         [["Balances can be transferred from any local bank account registered solely under the client’s own name."]]],

        ["Can a Standard Chartered Credit Card Funds Transfer pay off existing Standard Chartered debt?",
         [["No. A Credit Card Funds Transfer may not be used to pay off existing Standard Chartered loans or outstanding Credit Card payments."]]],

        ["What fees apply to a Standard Chartered Credit Card Funds Transfer?",
         ["The fees are:",
          B([["Early redemption fee: S$0, so there is no fee for paying off the balance early and no early repayment penalties."],
             ["Late payment charge: S$100."],
             ["There is no annual fee. A one-time, non-refundable processing fee applies based on the approved amount and tenure, from 0.90% with EIR from 2.06% p.a."]])]],

        ["What is the Effective Interest Rate on a Standard Chartered Credit Card Funds Transfer?",
         [["EIR is the true cost of borrowing. It accounts for the one-time processing fee, the monthly minimum repayments of 1% of principal, and full repayment of the balance in the final month of the tenure, which is why it is higher than the headline 0% interest figure."]]],

        ["What is the minimum monthly repayment on a Standard Chartered Credit Card Funds Transfer?",
         [["The minimum monthly repayment is 1% of the total remaining principal, or S$50, whichever is greater, plus any interest, fees and charges due. The remaining loan amount is payable at the end of the tenure."]]],

        ["What happens on a late Standard Chartered Credit Card Funds Transfer payment?",
         [["A flat S$100 late payment charge applies on the due date. If the minimum repayment is still not received in full by the next Statement Date, late interest at the EIR of 29.9% p.a. applies on all Credit Card Funds Transfer accounts from that Statement Date. Missing one payment means the promotional rate is reinstated once payment is made in full; missing two or more consecutive months may withdraw the promotional rate, with the 29.9% p.a. rate remaining even after catching up."]]],

        ["How can the 0% rate be lost on a Standard Chartered Credit Card Funds Transfer?",
         [["The 0% rate is lost if the full Credit Card Funds Transfer amount, including fees, charges and interest, is not repaid by the end of the tenure, or if the client is late on any Standard Chartered Credit Card, Personal Loan or Funds Transfer payment. In those cases the prevailing rate of 29.9% p.a. applies instead."]]],

        ["Do Standard Chartered Credit Card Funds Transfer balances earn rewards or cashback?",
         [["No. Credit Card Funds Transfer balances do not earn rewards or cashback."]]],

        ["Does a Standard Chartered Credit Card Funds Transfer affect the available credit limit?",
         [["Yes. It draws down the available credit limit, and as the transfer is repaid the available limit is restored."]]],

        ["How fast are Standard Chartered Credit Card Funds Transfer funds disbursed?",
         ["Disbursement timing depends on the receiving account and amount:",
          B([["To an existing Standard Chartered account: instantly upon approval, for any approved amount."],
             ["To another local bank account: 15 minutes."],
             ["If the approved amount exceeds S$200,000: funds arrive the next working day if approved before 2pm, or 2 working days after approval if approved after 2pm."],
             ["If the receiving bank is not on FAST: funds are disbursed the next working day instead."]])]],

        ["How are Standard Chartered Credit Card Funds Transfer repayments made?",
         [["Repayments are made via Online Banking, the SC Mobile app, AXS, ATMs, branches, or GIRO."]]],
      ]},

      { name: "Debt Consolidation Plan", url: U.dcp, qas: [
        ["What is the Standard Chartered Debt Consolidation Plan?",
         [["Debt Consolidation is a debt refinancing programme which offers a customer the option to consolidate all their unsecured credit facilities, such as credit cards and some types of unsecured loans, across financial institutions with one participating financial institution. Certain categories of unsecured loans are excluded, such as joint accounts, renovation loans, education loans, medical loans, and credit facilities granted for businesses or business purposes."]]],

        ["How does a Standard Chartered Debt Consolidation Plan help with debts?",
         ["Instead of paying varying amounts to different institutions each month, all debts become a single monthly payment. There are three main benefits:",
          B([["Easier tracking and planning, because consolidating into one payment makes the monthly amount owed clear."],
             ["Potentially lower interest rates compared to the outstanding debts, especially where these are mostly credit card debts, so less is usually paid overall."],
             ["Enforced financial discipline, because a single easily trackable monthly payment is easier to uphold."]])]],

        ["Who is eligible for a Standard Chartered Debt Consolidation Plan?",
         ["To be eligible for a Debt Consolidation Plan, an applicant must:",
          B([["Be a Singapore Citizen or Permanent Resident."],
             ["Earn between S$30,000 and below S$120,000 per annum."],
             ["Have total interest-bearing unsecured debt on all credit cards and unsecured credit facilities with financial institutions in Singapore that exceeds 12 times their monthly income."],
             ["Be aged 21 to 65 years old."]])]],

        ["What interest rates and tenures apply to a Standard Chartered Debt Consolidation Plan?",
         [["The Debt Consolidation Plan offers competitive interest rates from 3.48% p.a. (EIR from 6.26% p.a.), with a loan tenure from 3 to 10 years. The interest rate offered is based on the applicant’s personal credit profile as determined by the Bank and may differ from the published interest rate and from rates offered to other borrowers. Taking into account the joining fee of S$199, the EIR is 6.48% p.a. for a loan amount of S$20,000 at a 10-year tenure."]]],

        ["What credit card comes with a Standard Chartered Debt Consolidation Plan?",
         [["A Platinum Mastercard Credit Card is issued with a credit limit equivalent to the applicant’s monthly salary, as assessed by the Bank from the latest income document, for daily use. The credit limit is shared across all existing credit cards. The annual fee, which is S$192.60 including GST, is perpetually waived."]]],

        ["What savings does Standard Chartered illustrate for a Debt Consolidation Plan?",
         ["Standard Chartered gives an illustration assuming an outstanding balance of S$80,000:",
          B([["Current monthly payment to multiple banks: S$2,000 at prevailing interest rates, versus a new fixed instalment of S$1,187.95 with a Debt Consolidation Plan based on a 7-year loan."],
             ["Interest charges: S$21,520 per year calculated at EIR 26.9% p.a., versus S$2,826.89 per year calculated at 3.48% p.a. (EIR 6.48% p.a.)."],
             ["Interest savings with the Debt Consolidation Plan: S$18,693 per year."],
             ["The illustration assumes no new amounts are drawn down and that all instalment payments are paid by their due dates during the loan tenure."]])]],

        ["What documents are required for a Standard Chartered Debt Consolidation Plan?",
         ["All applicants must provide a copy of their NRIC front and back, a copy of their latest Credit Bureau report, the latest bank statements of their outstanding unsecured credit facilities such as credit cards, and the latest income documentation:",
          B([["Salaried employees and partial commission-based earners: latest computerised payslip, or latest 6 months’ CPF contribution history statement. The latest Income Tax Notice of Assessment should also be submitted to be considered for a higher loan amount."],
             ["100% commission-based earners: latest Income Tax Notice of Assessment, or latest 3 months’ commission statement from a single employer."],
             ["Self-employed: latest Income Tax Notice of Assessment."]])]],

        ["What fees and charges apply to a Standard Chartered Debt Consolidation Plan?",
         ["The fees and charges are:",
          B([["Joining fee, one time: S$199."],
             ["Early redemption fee: S$250 or 5% of the outstanding principal, whichever is higher."],
             ["Default interest, minimum EIR: 26.9% p.a. If payment is not received in full by the due date, finance charges are calculated on a daily basis at 0.074% from the respective transaction dates for all transactions to the date the payment is received."],
             ["Late payment charges: S$100, charged to the Debt Consolidation Plan facility and/or Platinum Mastercard Credit Card respectively if the minimum payment indicated on the statement is not received by the due date."],
             ["Annual fee for the Standard Chartered Platinum Mastercard Credit Card: S$192.60 including GST, perpetually waived."]])]],

        ["What is the early redemption fee on a Standard Chartered Debt Consolidation Plan?",
         [["The early redemption fee is S$250 or 5% of the outstanding principal, whichever is higher. For example, on a S$60,000 loan with an outstanding balance of S$10,000 before the loan tenure is up, an additional S$500 is payable, because the 5% fee is greater."]]],

        ["Can an existing Debt Consolidation Plan from another bank be refinanced with Standard Chartered?",
         [["Yes. An existing Debt Consolidation Plan held with another bank can be refinanced with Standard Chartered. A refinance promotion offering 6% cashback on refinancing an existing Debt Consolidation Plan is available, subject to terms and conditions."]]],
      ]},
    ],
  },

  // ============================================================= MORTGAGES
  {
    name: "Mortgages",
    subs: [
      { name: "Mortgage Range Overview", url: U.mortgages, qas: [
        ["Which mortgage products can a Standard Chartered Singapore client choose from?",
         ["Standard Chartered Singapore offers three mortgage products:",
          B([["MortgageOne, which offsets loan interest from interest earned on deposits."],
             ["Loan Repricing, which gives lower interest rates on an existing mortgage loan with the Bank."],
             ["Green Mortgage, which offers a pricing discount for properties with a valid BCA Green Mark rating."]])]],
      ]},

      { name: "MortgageOne", url: U.mortgageone, qas: [
        ["What is Standard Chartered MortgageOne?",
         ["MortgageOne reduces mortgage loan interest by offsetting it against interest earned on deposits. Its key features are:",
          B([["Reduced interest costs, by offsetting part of the home loan interest with eligible deposits."],
             ["Maintained flexibility, with full access to funds whenever needed."],
             ["Faster loan payoff, as more of the monthly instalment goes towards reducing the loan balance."]])]],

        ["What pricing and features apply to Standard Chartered MortgageOne?",
         ["MortgageOne pricing and features are:",
          B([["Interest rate: 3-month or 1-month Compounded SORA plus a margin for the Bank. Two-thirds of deposits enjoy the same rate as the mortgage loan, subject to a maximum of the loan principal outstanding. Remaining deposits earn an interest rate of 0.25% p.a."],
             ["Loan size: minimum SGD100,000."],
             ["Fees: no processing fee."],
             ["Type of property: private residential property only."]])]],

        ["What is the Standard Chartered MortgageOne National Day cashback promotion?",
         ["Applying directly with Standard Chartered from 3 to 10 August 2026 gives up to S$1,500 cashback on a MortgageOne loan:",
          B([["Loan amount from S$1 million to below S$1.5 million: up to S$1,000 cashback."],
             ["Loan amount of S$1.5 million and above: up to S$1,500 cashback."]])]],

        ["How does the Standard Chartered MortgageOne interest offset affect repayments?",
         [["Loan instalments are deducted monthly from the MortgageOne account. Any interest offset goes towards additional principal repayment for the month and keeps the overall instalment amount unchanged."]]],
      ]},

      { name: "Mortgage Facility Essentials", url: U.mortgageone, qas: [
        ["What is a Standard Chartered Mortgage Facility?",
         [["A Mortgage Facility is financing accorded for the purchase of property or the refinancing of an existing mortgage loan. Standard Chartered offers Mortgage Facilities for both Residential and Commercial Properties."]]],

        ["Which interest rate types does Standard Chartered offer on a Mortgage Facility?",
         ["Interest rate types generally fall into two broad categories, Fixed or Floating/Variable Rate, and the types offered may differ from time to time. Those offered include:",
          B([["Fixed Rate Packages."],
             ["Floating Rate Packages: 1-month Compounded Singapore Overnight Rate Average (SORA) and 3-month Compounded SORA. The 1-month and 3-month Compounded SORA for a given business day is published by 9am on the next business day on the Monetary Authority of Singapore website."],
             ["Fixed Deposit Rate (FDR). The tenure of the FDR offered may differ from time to time, and the actual tenure and corresponding interest rate is reflected in the Facility Letter for acceptance."]])]],

        ["How does the Fixed Deposit Rate work on a Standard Chartered Mortgage Facility?",
         [["The FDR offered may be different for other borrowers with different interest rate packages or FDR tenures. FDR is a floating or variable rate and is subject to change from time to time at the Bank’s sole discretion, depending on market conditions and any other factors impacting the benchmark of the FDR, with at least 30 days’ notice given."]]],

        ["What are the historical 36-month Fixed Deposit Rates on a Standard Chartered Mortgage Facility?",
         ["The historical rates for the 36-month Fixed Deposit Rate are:",
          B([["From 8 December 2022 to 27 April 2023: 1.52% per annum."],
             ["From 28 April 2023 to 27 April 2025: 2.77% per annum."],
             ["From 28 April 2025 to 21 August 2025: 1.77% per annum."],
             ["From 22 August 2025: 1.27% per annum."],
             ["These past trends may not reflect how high or low future interest rates may be."]])]],

        ["What fees and charges apply to a Standard Chartered Mortgage Facility?",
         ["A borrower is liable for the fees and charges relating to the Mortgage Facility, which may include but are not limited to:",
          B([["Legal fee, valuation fee and fire insurance premium relating to the purchase or mortgage of the property and the discharge of mortgage with the Bank."],
             ["Cancellation fee, payable on the amount of the Mortgage Facility cancelled or deemed cancelled after acceptance of the Facility Letter."],
             ["Partial or full redemption fee on the amount of the Mortgage Facility redeemed during the lock-in period, which is stipulated in the Facility Letter."],
             ["Repricing or restructuring fee for any change in the existing Mortgage Facility, for example a revision of interest rate package or loan tenure."],
             ["Fees for requesting copies of loan related documents or statements of account."],
             ["An annual administrative fee for taking up the insurance policy with the borrower’s own insurer."],
             ["Late fee and default fee on any late payments."]])]],

        ["How do repayments work on a Standard Chartered Mortgage Facility?",
         [["After the Mortgage Facility has been disbursed, the Bank notifies the borrower of the actual instalment payable and the effective date. The Bank also notifies the borrower when there is a revision of the monthly instalment. Any overdue instalment incurs interest at the default rate set out in the Facility Letter, plus the applicable late payment fee. The full Mortgage Facility Terms are published at ", ["av.sc.com", L.mortgageTerms], "."]]],

        ["What is the Residential Property Loan Sheet from Standard Chartered?",
         [["After a discussion about financing needs, Standard Chartered provides a copy of the Residential Property Loan Sheet, which contains key features of the Mortgage Facility, loan tenure, interest rate, monthly instalment, and information on the impact to payments in the event of rising interest rates. Borrowers are also encouraged to refer to The Association of Banks in Singapore’s Consumer Guide on Home Loans, available on the MoneySENSE and ABS websites, before committing to a loan."]]],
      ]},

      { name: "SORA Pricing Package", url: U.sora, qas: [
        ["What is SORA?",
         [["The Singapore Overnight Rate Average (SORA) is the volume-weighted average rate of borrowing transactions in the unsecured overnight interbank SGD cash market in Singapore between 8am and 6.15pm."]]],

        ["How is SORA computed?",
         [["On each business day in Singapore, reporting banks provide data on all eligible transactions traded and booked in the window between 8am and 6.15pm, both timings inclusive. The Monetary Authority of Singapore conducts thorough data validation checks and computes SORA by taking the volume-weighted average rate of all eligible transactions. SORA is then published on the MAS website the next business day at 9am."]]],

        ["What is the Standard Chartered 3M Compounded SORA promotional pricing package?",
         ["The 3M Compounded SORA promotional pricing package is:",
          B([["Year 1: 3M Compounded SORA plus 1.00% p.a."],
             ["Year 2: 3M Compounded SORA plus 1.00% p.a."],
             ["Year 3: 3M Compounded SORA plus 1.00% p.a."],
             ["Thereafter: 3M Compounded SORA plus 1.00% p.a."],
             ["Lock-in period: 2 years. Minimum loan amount: S$100,000."],
             ["The package is valid from 27 February 2021 inclusive and applies to residential property loans only."]])]],

        ["What is 3M Compounded SORA?",
         [["The 3M Compounded SORA means the 3-Month Compounded Singapore Overnight Rate Average. The 3-Month Compounded Singapore Overnight Rate Average for a given business day is published by 9am on the next business day on the Monetary Authority of Singapore website, or on any other website designated by MAS. It is computed by compounding the published SORA rate over the historical 3-month period and has been published by MAS since 5 August 2020."]]],

        ["How are Standard Chartered SORA pricing and interest rates determined?",
         [["The pricing is first determined on the first business day of the loan disbursement month. Thereafter, repricing of the 3M Compounded SORA is done every three months, on the first business day of the month. The Bank notifies the borrower of any revisions made to monthly instalments when they are due."]]],

        ["Which properties are eligible for a Standard Chartered SORA package?",
         [["Only Private Residential properties and HDB flats, whether completed or Building Under Construction, are eligible for the SORA package, used for a new purchase or refinancing."]]],

        ["Can an existing Standard Chartered loan be changed to a SORA interest rate package?",
         [["Yes, once the existing loan is out of the lock-in period. All repricing applications are subject to approval and repricing fee charges. Repricing can be requested via Online Banking or the SC Mobile app under Help & Services."]]],
      ]},

      { name: "Green Mortgage", url: U.green, qas: [
        ["What is the Standard Chartered Green Mortgage Promotion?",
         ["The Green Mortgage Promotion gives a pricing discount on a mortgage loan when the following conditions are met:",
          B([["The mortgage property must have a valid Building and Construction Authority (BCA) Green Mark rating of Green Mark Gold, Green Mark GoldPlus, Green Mark Platinum, or Green Mark Super Low Energy (SLE), which includes Green Mark GoldPlus SLE and Green Mark Platinum SLE, and which is within 3 years from the certification year."],
             ["The applicant must apply for a new mortgage loan or reprice their existing loan with the Bank."],
             ["The applicant must take up a floating rate package, which is a SORA-Pegged package."]])]],

        ["When does the Standard Chartered Green Mortgage Promotion run, and who is eligible?",
         [["The Green Mortgage Promotion is available until 31 July 2026, both dates inclusive, to all new and existing customers of Standard Chartered Bank (Singapore) Limited who are individuals. The applicant must submit an application during the Promotion Period, either for a new Eligible Mortgage Facility or to reprice an existing Mortgage Facility to an Eligible Mortgage Facility. The BCA Green Mark Rating should be stated in the application form, or on the call with the Bank’s mortgage repricing unit for a Repricing Application."]]],

        ["What is the BCA Green Mark Rating?",
         [["The BCA Green Mark certification scheme was launched in 2005 and is a green building rating system designed to evaluate a building’s environmental impact and performance from design to operational stage. The scheme is refreshed every few years, the latest being the BCA Green Mark 2021 certification scheme, which encourages the industry and professionals to collaborate and develop green building solutions to raise energy efficiency and attain sustainability outcomes outlined in the Singapore Green Building Masterplan. One of the targets under that Masterplan is to have at least 80% of buildings in Singapore, by Gross Floor Area, green by 2030."]]],

        ["How is the BCA Green Mark rating of a mortgaged property checked?",
         [["The BCA Green Mark rating of a property can be checked on the SLEB Green Mark Buildings Directory at ", ["sleb.sg", L.greenMarkDir], "."]]],

        ["Can more than one property qualify for the Standard Chartered Green Mortgage Promotion?",
         [["Yes, provided the properties meet the eligibility criteria."]]],

        ["How does Standard Chartered define what is “green”?",
         [["The Bank has a Green and Social Sustainable Finance Product Framework which defines eligibility for a sustainability label for all products across the Bank. It was co-authored with Sustainalytics, a leading provider of ESG data and verification."]]],

        ["How else does Standard Chartered support sustainability?",
         [["The Bank has various products that support sustainability by generating social and environmental benefits. For Green Mortgages, the Bank encourages customers to purchase green properties which have lower carbon footprints and consume less energy. Other products that support sustainability are ", ["Sustainable Time Deposits", L.sustainableTd], " and ESG Unit Trusts."]]],

        ["What is the scale of Standard Chartered’s Sustainable Finance delivery?",
         ["Standard Chartered reports the following on its Sustainable Finance delivery:",
          B([["Social: over 885,000 microfinance loans enabled."],
             ["Environmental: 1.4 million tonnes of CO2 emissions avoided."],
             ["Financial: nearly 20,000 SME loans disbursed."],
             ["Geographical: 70% of total Sustainable Finance assets are located in emerging markets, and 84% of those are in Asia, Africa and the Middle East."]])]],

        ["What fees and charges apply to a Standard Chartered Green Mortgage?",
         ["The fees and charges are:",
          B([["Late payment fee: S$50 for SGD loans."],
             ["Default rate: 5% above the SGD Prime lending rate on the overdue instalment amount for SGD loans. The Prime lending rate is a minimum indicative interest rate charged by the Bank and is usually used as a benchmark for loans."],
             ["The full list of fees and charges is published in the ", ["Standard Chartered Singapore Pricing Guide", L.pricingGuide], "."]])]],
      ]},

      { name: "Loan Repricing — Eligibility and Documents", url: U.repricing, qas: [
        ["What is Standard Chartered mortgage loan repricing?",
         ["Repricing gets more out of an existing mortgage by moving to a lower interest rate and monthly instalment. Its benefits are:",
          B([["A lower interest rate and monthly instalment."],
             ["Savings on the legal and valuation fees typically incurred when refinancing with another bank."],
             ["A fast application process, so potential savings start sooner."],
             ["A hassle-free experience with minimal documents needed."]])]],

        ["Who is eligible for Standard Chartered mortgage loan repricing?",
         [["An applicant should have an outstanding loan balance of S$100,000 and a remaining loan tenure of at least 5 years."]]],

        ["What documents are required to reprice an owner-occupied property loan with Standard Chartered?",
         ["For SingPass holders applying with MyInfo on an owner-occupied property, the IRAS Tax Portal Page is required, obtained as follows:",
          B([["Step 1: Log in to myTax Portal on the IRAS website using SingPass."],
             ["Step 2: Under Property, click View Property Summary."],
             ["Step 3: Click Save as PDF or Print to save a PDF copy."],
             ["Foreign applicants must additionally provide a passport, and a utility bill or bank statement reflecting the subject property address as their mailing address."]])]],

        ["How does Standard Chartered assess a repricing application for an investment property or a property with an equity loan?",
         ["There are 2 ways to assess such a repricing application:",
          B([["Debt Reduction Plan (DRP): the loan can be repriced with fewer documents, without going through the full Total Debt Servicing Ratio assessment, if the borrower repays 3% of their total outstanding loan amount upfront or in monthly instalments over 3 years."],
             ["Total Debt Servicing Ratio (TDSR) assessment."]])]],

        ["What documents are required to reprice an investment property loan with Standard Chartered?",
         ["Requirements for an investment property only are:",
          B([["Under the Debt Reduction Plan: no document is required for SingPass holders, as all personal and declaration data is retrieved using MyInfo. Foreign applicants must provide a passport."],
             ["Under the Total Debt Servicing Ratio assessment: the latest payslip for salaried borrowers only. Foreign applicants must additionally provide a passport."]])]],

        ["What CPF documents are required to reprice a Standard Chartered property loan with an equity loan?",
         ["For an investment property with an equity loan, or an owner-occupied property with an equity loan, CPF withdrawal details are required on both principal and accrued interest and on monthly instalment, obtained as follows:",
          B([["Log in to my CPF Online Services on the CPF website using SingPass."],
             ["Under my cpf, go to My dashboards, then select Home ownership."],
             ["Select Principal amount and accrued interest, or Monthly CPF deduction, as applicable."],
             ["Click the PDF icon at the top right corner to save a PDF copy."],
             ["For an owner-occupied property with an equity loan, the IRAS Tax Portal Page must also be provided."]])]],

        ["What fees and charges apply to Standard Chartered mortgage loan repricing?",
         ["The fees and charges are:",
          B([["Repricing fee: an administrative fee of S$1,000."],
             ["Late payment fee: S$50 for SGD loans."],
             ["Default rate: 5% above the SGD Prime lending rate on the overdue instalment amount for SGD loans."]])]],
      ]},

      { name: "Loan Repricing — Rates and Requests", url: U.repricing, qas: [
        ["How is a Standard Chartered mortgage repricing requested?",
         ["Repricing is requested through Online Banking or the SC Mobile app as follows:",
          B([["Step 1: Log in to Online Banking or the SC Mobile app."],
             ["Step 2: Under “Help & Services”, select “Mortgage Management”."],
             ["Step 3: Select “Mortgage Repricing”."],
             ["Step 4: Select the loan account."],
             ["Step 5: Review the request and click “Next” to submit it."]])]],

        ["Why should a Standard Chartered mortgage be repriced instead of refinanced?",
         [["Repricing allows a borrower to enjoy the new loan package within a month, while refinancing typically takes about 3 months, so interest savings start earlier with repricing. Repricing could involve a repricing admin fee of about S$1,000, whereas refinancing requires legal and valuation fees usually above S$2,000. Engaging a law firm and valuation company is not required when repricing, which saves time and money."]]],

        ["How often do Standard Chartered SORA package rates change, and how is a borrower notified?",
         [["For a mortgage loan based on the 3-month Compounded SORA, the interest rate is repriced on the 1st business day every three months. The borrower receives a letter from the Bank or an eAdvice in the SC Online Banking inbox every 3 months updating them of any rate changes, and the monthly instalment change only takes effect 2 months later, which allows time to change any existing standing instructions. For example, if there is a change of interest rate on 1 March 2022, the first notice arrives in March, but the change in monthly instalment amount only takes effect on 1 May 2022. Any subsequent notices serve only as a reminder. The Base Rate Change Letter contains the full information on the effective date of the interest rate and monthly instalment amount change."]]],

        ["How can a borrower check and change the CPF amount paid towards a Standard Chartered mortgage?",
         [["The amount can be checked by logging on to Online Banking, the SC Mobile app, or the CPF website with SingPass. To make changes to monthly CPF payment arrangements, log in to the CPF website and go to “Tools and Services”, then “Forms and E-applications”, then “Manage CPF usage for your home”, then “Apply Online”. It typically takes 1 working day for the CPF board to effect the revised amount."]]],
      ]},

      { name: "Mortgage Loan Redemption", url: U.repricing, qas: [
        ["How is a partial redemption of a Standard Chartered mortgage loan made?",
         ["A mortgage loan can be redeemed in part as follows:",
          B([["Serve the Bank a 1-month notice in writing, or pay 1 month’s interest in lieu of notice."],
             ["The partial prepayment must be a minimum of S$10,000."],
             ["A 1.50% partial redemption fee is payable on the amount of the Mortgage Facilities prepaid during the Lock-In Period."],
             ["A redemption form must be retrieved to proceed."]])]],

        ["Can CPF be used for a partial redemption of a Standard Chartered mortgage loan?",
         [["Yes. A loan can be redeemed in part using CPF, and there is no requirement to serve a 1-month notice in writing. The instruction is submitted via the CPF website using SingPass, or at any CPF branch office."]]],

        ["How is a full redemption of a Standard Chartered mortgage loan made?",
         ["A mortgage loan can be redeemed in full as follows:",
          B([["Serve the Bank a 2 months’ notice in writing, or pay 2 months’ interest in lieu of notice."],
             ["A 1.50% full redemption fee is payable on the amount of the Mortgage Facilities redeemed during the Lock-In Period."],
             ["CPF may be used to redeem the loan in full, but 2 months’ notice in writing must still be served, otherwise 2 months’ interest in lieu of notice is incurred."],
             ["A redemption form must be retrieved to proceed."]])]],

        ["Who appoints the lawyer for a Standard Chartered mortgage full redemption?",
         [["Standard Chartered appoints a law firm to process the full redemption if the borrower has not done so. A borrower who prefers to appoint their own law firm should note that the choice of law firm is subject to the Bank’s approval. If the appointed lawyers are not on the Bank’s panel list, additional costs may be incurred, as the borrower has to bear the fees of both their own and the Bank’s lawyers for the full redemption."]]],

        ["What costs are involved in a Standard Chartered mortgage full redemption?",
         [["On redeeming the Mortgage Facilities, the borrower pays all fees relating to the discharge of the Mortgage Property, such as legal fees including the Bank’s legal fees and discharge of mortgage. These fees are billed directly by the chosen law firm. Legal fees for the discharge of property typically range from S$900 to S$1,200. This range is indicative only, and the law firm should be checked with directly for actual costs before proceeding."]]],

        ["When is the title deed received after redeeming a Standard Chartered mortgage loan?",
         [["The Singapore Land Authority typically takes 3 weeks to 1 month to release the title deed to the lawyer. The lawyer should then contact the borrower for collection within a week."]]],
      ]},

      { name: "Mortgage Fire Insurance", url: U.repricing, qas: [
        ["What is fire insurance on a Standard Chartered mortgaged property?",
         [["Fire insurance for a residential property provides coverage against the loss or damage to the insured property as a result of fire and extraneous perils as stated in the policy. The policy is commonly referred to as a Mortgagee Interest Policy (MIP)."]]],

        ["Why is a Mortgagee Interest Policy required on a Standard Chartered mortgage?",
         [["A Mortgagee Interest Policy is required because, if there is a failure to service the loan as a result of damage to the mortgaged home, the Bank can make a claim against the insurer on the MIP. The homeowner remains liable for the outstanding property loan amount to the MIP Insurer."]]],

        ["Does a Mortgagee Interest Policy duplicate a condominium’s existing fire insurance?",
         [["No. Under the Land Titles (Strata) Act (Chapter 158), the Management Corporation of the property must insure a “Damage Policy”, a fire insurance policy which ensures that the entire development, including its common property, is protected. As the Damage Policy and the MIP serve different protection needs and purposes, there is no duplication of protection."]]],

        ["What is Sum Insured on a Standard Chartered mortgaged property?",
         [["Sum insured is the reinstatement cost needed to rebuild a home after a fire, and not the market value of the property. The land value is not included, because in the event of a total loss the land will still be present."]]],

        ["How is the Sum Insured Value determined on a Standard Chartered mortgaged property?",
         [["The Sum Insured Value is determined by the Bank. For private apartments with strata-title, the SIV is based on the reinstatement value or the outstanding loan amount, whichever is lower. For private apartments without strata-title and for landed property, the SIV is based on the reinstatement value."]]],

        ["Can a borrower use their own insurance company for Standard Chartered mortgage fire insurance?",
         [["Yes. However, any request for a change in insurance company is subject to the Bank’s approval, and the Bank must be notified of the intention to use an alternative insurance company."]]],

        ["What is required for a self-arranged fire insurance policy on a Standard Chartered mortgage?",
         ["The following documents must be provided to the Bank 5 business days before the expiry of the insurance policy:",
          B([["The original copy of a signed letter of undertaking for the Bank’s safe-keeping."],
             ["The original fire insurance policy stating the Bank as the Mortgagee."],
             ["A non-cancellation clause in the policy stating that the Bank must be informed by the insurer prior to any cancellation or any material changes proposed to the policy."],
             ["The original receipt showing the premium was paid."],
             ["If these documents are not furnished before the expiry of the insurance policy, the Bank will instruct Allianz Global Corporate & Specialty SE Singapore Branch to effect the fire insurance. An upfront administration fee of S$100 applies, and thereafter an annual administration fee of S$100 is charged."]])]],
      ]},
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Borrow FAQ",
  docTitle: "SCB Bank_Borrow FAQ",
  description: "Standard Chartered Singapore Borrow FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
console.log("categories:", CATEGORIES.length, "subcategories:", CATEGORIES.reduce((n, c) => n + c.subs.length, 0));
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Borrow_FAQ.docx", buf);
  console.log("written", buf.length);
});
