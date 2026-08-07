const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

// The Pricing Guide is a single page; every subcategory cites the same source.
const PG = "https://www.sc.com/sg/pricing-guide/";

const L = {
  priorityTnc: "https://www.sc.com/sg/terms-and-conditions/priority/",
  ccft: "https://www.sc.com/sg/borrow/loans/balance-transfer/",
};

const PRIME_NOTE = "The Prime lending rate is a minimum indicative interest rate charged by the Bank and is usually used as a benchmark for loans.";

const CATEGORIES = [
  // ====================================================== DEPOSIT ACCOUNTS
  {
    name: "Deposit Accounts",
    subs: [
      { name: "Savings Accounts", url: PG, qas: [
        ["What fees apply to the Standard Chartered e$aver and Unlimited$aver savings accounts?",
         ["The Personal Banking fees are as follows, and all of them are waived for Priority Banking clients:",
          B([["e$aver fall-below fee: S$5 per month if the average daily balance for the month falls below S$1,000."],
             ["e$aver counter transaction fee: S$5 per cash withdrawal."],
             ["Unlimited$aver fall-below fee: S$5 per month if the average daily balance for the month falls below S$3,000."]])]],

        ["What fees apply to the Standard Chartered MyWay and JumpStart savings accounts?",
         ["The Personal Banking fees are as follows, and all of them are waived for Priority Banking clients:",
          B([["MyWay fall-below fee: S$10 per month if the average daily balance for the month falls below S$50,000."],
             ["JumpStart fall-below fee: waived."],
             ["JumpStart debit card annual fee: waived."]])]],

        ["What fees apply to the Standard Chartered Basic Bank Account?",
         ["The Personal Banking fees are as follows, and all of them are waived for Priority Banking clients:",
          B([["Monthly service charge: S$2 regardless of balance, waived for beneficiaries under the MSF Public Assistance Scheme or Special Grant Scheme."],
             ["Early account closure fee within 6 months: S$30."],
             ["Passbook replacement fee: S$30."]])]],

        ["What fees apply to the Standard Chartered Savings Account, also known as the Supersave account?",
         ["The Personal Banking fees are as follows, and all of them are waived for Priority Banking clients:",
          B([["Fall-below fee: S$5 per month if the average daily balance for the month falls below S$2,000."],
             ["Monthly service charge: S$5."],
             ["Early account closure fee within 6 months: S$30."],
             ["Passbook replacement fee: S$30."]])]],
      ]},

      { name: "Current Accounts", url: PG, qas: [
        ["What fees apply to the Standard Chartered Wealth $aver current account?",
         ["The Personal Banking fees are as follows, and all except the unarranged overdraft are waived for Priority Banking clients:",
          B([["Fall-below fee: S$5 per month if the average daily balance for the month falls below S$3,000, from the second month after account opening."],
             ["Monthly service charge: S$2."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30."],
             ["Debit card annual fee: S$20 inclusive of GST."],
             ["Cheque book: S$10 per cheque book requested."]])]],

        ["What fees apply to the Standard Chartered Bonus$aver account?",
         ["The Bonus$aver fees are:",
          B([["Fall-below fee: S$5 per month if the average daily balance for the month falls below S$3,000, waived for Priority Banking."],
             ["Early account closure fee within 6 months: S$30, waived for Priority Banking."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Debit card annual fee: waived."],
             ["Credit card annual fee: S$218 inclusive of GST, free for the first 2 years, for both Personal Banking and Priority Banking."],
             ["Cheque book: S$10 per cheque book requested, waived for Priority Banking."]])]],

        ["What fees apply to the Standard Chartered XtraSaver account?",
         ["The Personal Banking fees are as follows, and all except the unarranged overdraft are waived for Priority Banking clients:",
          B([["Fall-below fee: S$5 per month if the average daily balance for the month falls below S$3,000, from the second month after account opening."],
             ["Monthly service charge: S$2."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30."],
             ["Debit card annual fee: waived."],
             ["Cheque book: S$10 per cheque book requested."]])]],

        ["What fees apply to the Standard Chartered SuperSalary account?",
         ["The Personal Banking fees are as follows, and all except the unarranged overdraft are waived for Priority Banking clients:",
          B([["Non-inward credit fee: S$5 charged to the account, from the fourth month after account opening, if the amount credited for the month is below S$1,500."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30."],
             ["Cheque book: S$10 per cheque book requested."]])]],

        ["What fees apply to the Standard Chartered Cheque and Save account?",
         ["The Personal Banking fees are as follows, and all except the unarranged overdraft are waived for Priority Banking clients:",
          B([["Fall-below fee: S$7.50 per month if the average daily balance for the month falls below S$5,000."],
             ["Monthly service charge: S$2."],
             ["Unarranged overdraft: Prime + 5%, minimum S$5."],
             ["Early account closure fee within 6 months: S$30 for personal accounts and S$50 for corporate accounts."],
             ["Cheque book: S$10 per cheque book requested."]])]],

        ["What fees apply to the Standard Chartered OneAccount, MortgageOne and Wealth Cash Account?",
         ["The fees are:",
          B([["OneAccount, which is available for existing accounts only: fall-below fee of S$5 per month if the average daily balance falls below S$3,000; monthly service charge of S$2; unarranged overdraft at Prime + 5%, minimum S$5; and cheque book at S$10 per book. All except the unarranged overdraft are waived for Priority Banking."],
             ["MortgageOne: unarranged overdraft at Prime + 5%, minimum S$5, and cheque book at S$10 per book, waived for Priority Banking."],
             ["Wealth Cash Account: unarranged overdraft at Prime + 5%, minimum S$5."],
             [PRIME_NOTE]])]],
      ]},

      { name: "Foreign Currency Accounts", url: PG, qas: [
        ["What fees apply to the Standard Chartered USD$aver foreign currency savings account?",
         ["The Personal Banking fees are:",
          B([["Fall-below fee: US$10 per month if the average daily balance for the month falls below US$10,000, waived for Priority Banking."],
             ["Unarranged overdraft: Prime + 5%, minimum US$5."],
             ["Early account closure fee within 6 months: US$20, waived for Priority Banking."],
             ["Debit card overseas transaction fee: 1% Mastercard fee for transactions made at overseas merchants."]])]],

        ["What fall-below fees apply to the Standard Chartered FCY$aver account by currency?",
         ["The FCY$aver fall-below fee is charged monthly if the average daily balance falls below the stated threshold, and is waived for Priority Banking clients:",
          B([["AUD: A$5 below A$2,000. CAD: C$5 below C$2,000. NZD: NZ$5 below NZ$2,000. USD: US$5 below US$2,000."],
             ["CHF: 5 francs below 2,000 francs. EUR: €3 below €1,000. GBP: £2 below £1,000."],
             ["CNH: ¥40 below ¥15,500. HKD: HK$30 below HK$15,000. JPY: ¥500 below ¥200,000."],
             ["FCY$aver early account closure fee: US$20, waived for Priority Banking."]])]],

        ["What fees apply to the Standard Chartered USD High foreign currency current account?",
         ["The Personal Banking fees are:",
          B([["Fall-below fee: US$10 per month if the average daily balance for the month falls below US$10,000, waived for Priority Banking."],
             ["Unarranged overdraft: Prime + 5%, minimum US$5."],
             ["Early account closure fee within 6 months: US$20, waived for Priority Banking."],
             ["Debit card overseas transaction fee: 1% Mastercard fee for transactions made at overseas merchants."]])]],

        ["What fall-below fees apply to the Standard Chartered foreign currency Current Account by currency?",
         ["The foreign currency Current Account fall-below fee is charged monthly if the average daily balance falls below the stated threshold, and is waived for Priority Banking clients:",
          B([["AUD: A$30 below A$5,000. NZD: NZ$30 below NZ$5,000. USD: US$30 below US$5,000."],
             ["CHF: 30 francs below 5,000 francs. EUR: €20 below €3,000. GBP: £15 below £2,000. HKD: HK$200 below HK$30,000."],
             ["Other charges: unarranged overdraft at Prime + 5%, minimum US$5; early account closure fee within 6 months of US$20 personal or US$30 corporate; and USD cheque book at US$10 per book."]])]],

        ["What fees apply to the Standard Chartered foreign currency Cheque and Save account?",
         ["The fall-below fee is charged monthly if the average daily balance falls below the stated threshold, and all fees are waived for Priority Banking clients:",
          B([["Fall-below fee: A$10 below A$5,000; C$10 below C$5,000; 10 francs below 5,000 francs; €5 below €3,000; £5 below £2,000; HK$75 below HK$30,000; ¥1,000 below ¥500,000; NZ$10 below NZ$5,000; US$10 below US$5,000."],
             ["Monthly service charge: A$2, C$2, 2 francs, €1, £1, HK$15, ¥200, NZ$2 or US$2 depending on currency."],
             ["Unarranged overdraft: Prime + 5%, minimum US$5."],
             ["Early account closure fee within 6 months: US$20 personal or US$30 corporate."],
             ["USD cheque book: US$10 per cheque book requested."]])]],

        ["What unarranged overdraft rate applies to a Standard Chartered foreign currency Wealth Cash Account?",
         [["The foreign currency Wealth Cash Account, available in USD, AUD, EUR, GBP, HKD, JPY, CHF, NZD, CAD, NOK, SEK, ZAR, DKK and CNH, carries an unarranged overdraft rate of Prime + 2%, with no minimum. " + PRIME_NOTE]]],
      ]},

      { name: "Time Deposits", url: PG, qas: [
        ["What penalty applies to premature upliftment of a Standard Chartered Foreign Currency Time Deposit?",
         ["The total penalty is the interest accrued plus a handling fee of USD25 plus the replacement cost, calculated as follows:",
          B([["Replacement Cost = (Current Market Day Rate – Contract Rate) x Remaining Days to Maturity."],
             ["The Replacement Cost can be zero if the Current Market Rate at the date of the premature upliftment is less than the Contract Rate."],
             ["Current Market Rate refers to the current interest rate, based on market conditions for the applicable currency, at the date of the premature upliftment."],
             ["Contract Rate refers to the interest rate that would have been payable if the Time Deposit had been held until maturity."]])]],
      ]},
    ],
  },

  // ================================================ WEALTH AND INVESTMENTS
  {
    name: "Wealth and Investments",
    subs: [
      { name: "Wealth Solutions — Pricing Basis", url: PG, qas: [
        ["How should the Standard Chartered Pricing Guide for wealth solutions be read?",
         ["The Pricing Guide sets out pricing information relating to products and services the Bank may make available. In addition:",
          B([["For trading and investments, certain fees and charges may be payable, set out in the relevant investment and insurance product documents."],
             ["The Bank may levy a transaction fee or charge, or earn a commission, when executing certain transactions, and may earn a mark-up or receive a monetary benefit when a client transacts in certain products."],
             ["The pricing information may not include all third-party charges, which are in addition and levied separately."],
             ["Where a product or service requires exceptional handling, the Bank has the right to levy extra fees for the additional work required."],
             ["Transactions involving certain assets and investment types may be subject to different pricing, and the Bank will inform the client if pricing for a transaction exceeds what is stated."],
             ["Goods and services tax will be charged where applicable."]])]],

        ["How does Standard Chartered quote prices on wealth transactions?",
         [["Product costs and charges are associated with manufacturing and managing a financial instrument, are incurred or specified by the manufacturer, and are either applied to that instrument or taken into account in its overall pricing. The Bank generally quotes an all-in price for a transaction, which includes such costs and charges and any sales and trading mark-ups. If the Bank is then able to execute the transaction at a better interbank price, it will endeavour to pass on the benefit of that price improvement. Where the Bank determines it is not operationally feasible to pass on the benefit, which applies to FX Spot and FX over-the-counter derivatives contracts, the Bank is entitled to retain all of that benefit."]]],

        ["In what capacity does Standard Chartered act on wealth transactions?",
         [["The Bank acts as principal when providing each type of product, investment or service described, except that it acts as agent for the sale or purchase of equities for a client’s account and for the sale or purchase of investment funds. With respect to securitised products, any product issuer, provider, bookrunner or arranger that has “Standard Chartered” in its name is an affiliate of Standard Chartered Bank. A contract note for a transaction may contain further disclosure of applicable monetary benefits, and where a contract note and the Pricing Guide are inconsistent, the contract note prevails."]]],
      ]},

      { name: "Cash Equities — Brokerage and Custody", url: PG, qas: [
        ["What online brokerage rates apply to Standard Chartered cash equities trading?",
         ["Online brokerage and custody rates are:",
          B([["Priority Banking clients: 0.18% on SGX and 0.20% on all other markets, with a custody fee of 0 and no minimum brokerage amount."],
             ["Personal Banking clients: 0.20% on SGX and 0.25% on all other markets, with a custody fee of 0."],
             ["Personal Banking minimum brokerage amount: 10 AUD/CHF/EUR/GBP/SGD/USD where shares are traded in those currencies, 100 HKD where traded in HKD, and 1,000 JPY where traded in JPY."],
             ["The minimum brokerage amount is based on the currency of the shares traded, regardless of the exchange the shares are listed on. GST is imposed where applicable."]])]],

        ["What offline brokerage rates apply to Standard Chartered cash equities trading?",
         ["Offline brokerage and custody rates are:",
          B([["Priority Banking clients: 0.33% on SGX and 0.35% on all other markets, with a custody fee of 0."],
             ["Personal Banking clients: 0.35% on SGX and 0.40% on all other markets, with a custody fee of 0."],
             ["Minimum brokerage amount for both segments: 100 AUD/CHF/EUR/GBP/SGD/USD where shares are traded in those currencies, 550 HKD where traded in HKD, and 10,000 JPY where traded in JPY."],
             ["The minimum brokerage amount is based on the currency of the shares traded, regardless of the exchange the shares are listed on. GST is imposed where applicable."]])]],

        ["What are Standard Chartered cash equities trades subject to?",
         [["All trades are subject to brokerage fees levied by Standard Chartered Bank (Singapore) Limited and market fees levied by the respective exchange. GST is imposed where applicable."]]],
      ]},

      { name: "Cash Equities — Market and Transfer Fees", url: PG, qas: [
        ["Which exchanges charge no market fees on Standard Chartered cash equities trades?",
         [["No market fee is charged on the Australian SE (ASX), Deutsche Boerse (XETR), Tokyo SE (TSE), NYSE EN Amsterdam (AMS) and SIX Swiss Exchange (SWX)."]]],

        ["What market fees apply on Standard Chartered trades in Singapore, Hong Kong and the United States?",
         ["The market fees are:",
          B([["Singapore SE (SGX) and Singapore SE Odd Lot (SGXO): total market fees of 0.04%, comprising an SGX Clearing Fee of 0.0325% and an SGX Trading Access Fee of 0.0075%."],
             ["Hong Kong SE (HKG): Stamp Duty of 0.1% rounded up to the nearest dollar effective 17 November 2023, Transaction Levy of 0.0027%, Trading Fee of 0.00565%, and Financial Reporting Council Transaction Levy of 0.00015%."],
             ["United States exchanges, covering NYSE AMEX, NASDAQ, NASDAQ OTC, NYSE ARCA, Cboe Global Markets and the New York Stock Exchange: Securities and Exchange Commission Fee of 0.00206% for sell trades only, effective 2 April 2026."],
             ["ADR fees range from USD 0.005 to 0.05 and are charged for applicable corporate action events."]])]],

        ["What market fees apply on Standard Chartered trades in France and the United Kingdom?",
         ["The market fees are:",
          B([["NYSE EN Paris (PAR): French Financial Transaction Tax of 0.4% on buy trades only, effective 1 April 2025."],
             ["London Stock Exchange (LSE): Stamp Duty of 0.5% on buy trades for GB ISIN shares only, Stamp Duty of 1.00% on buy trades for IE ISIN shares only, and a levy of 1.5 GBP for any transaction over GBP 10,000 on both buy and sell trades."]])]],

        ["What share transfer fees apply to SGX-listed shares on SC Online Trading?",
         ["The fees for SGX-listed shares, with prevailing GST applicable, are:",
          B([["Inward transfer into SC Online Trading with no change in beneficial ownership: S$10.00 per counter, per transfer."],
             ["Inward or outward transfer with a change in beneficial ownership: S$10.00 per counter, per 1,000 shares, per transfer, up to a maximum of S$100.00 per counter."],
             ["Internal transfer within SC Online Trading with no change in beneficial ownership: free."],
             ["Outward transfer from SC Online Trading with no change in beneficial ownership: S$10.00 per counter, per transfer."]])]],

        ["What share transfer fees apply to foreign-listed shares on SC Online Trading?",
         ["Inward and internal transfers of foreign-listed shares are free. Outward transfers are charged based on the trading currency of the shares, regardless of the exchange they are listed on, on a per counter, per transfer basis:",
          B([["AUD 55, CHF 38, GBP 30, USD 39, EUR 34, HKD 305 and JPY 4,307."],
             ["Additional charges apply if shares are physical certificates."],
             ["Change of beneficial ownership is not allowed for UK, France and Hong Kong listed shares."]])]],

        ["Which share transfers are permitted on SC Online Trading?",
         ["For inward, internal and outward transfers, the permitted combinations are:",
          B([["Client A to Client A: permitted."],
             ["Client A to Client A and/or B: permitted."],
             ["Client A to Client B: not permitted."],
             ["Client A and/or B to Client A or Client B only: not permitted."],
             ["Hong Kong, France and UK share transfers involving a change in beneficial ownership are not permitted."]])]],

        ["What fee applies to issuance of physical share certificates through Standard Chartered?",
         [["Issuance of physical share certificates costs USD200 per certificate, on top of any miscellaneous pass-through charges. Issuance for delisted shares is subject to the discretion of the company and the share registrar. Transferring in and conversion of physical share certificates is not supported, and exceptional handling in certain cases is at the discretion of the Bank. GST is imposed where applicable."]]],
      ]},

      { name: "Capital Market Products and Funds", url: PG, qas: [
        ["What monetary benefits does Standard Chartered receive on capital market products?",
         ["The Bank receives the following benefits:",
          B([["Fixed Income, including Bonds, Notes, Treasury Bills and Certificates of Deposits: up to 3.00% of notional amount."],
             ["Primary Market Rebates: up to 1.00% of notional amount, received from the issuer."],
             ["Foreign Exchange Spot and Commodity Spot on unallocated precious metals such as XAU and XAG: up to 5.00% of notional amount."],
             ["Structured Products, including Equity Linked, Rate Linked, Currency Linked and Commodity Linked Structured Notes: up to 5.00% of notional amount."]])]],

        ["What fees does Standard Chartered receive on mutual funds and hedge funds?",
         ["The Bank receives the following fees:",
          B([["Investment Funds on buy transactions: up to 5.00% of the investment amount, dependent on the fund’s fees and charges as prescribed in its prospectus."],
             ["Trailer fee for mutual funds: up to 70% of the fund’s annual management fee as prescribed in the fund’s prospectus."],
             ["Trailer fee for hedge funds: up to 1% of the invested amount, derived from the management fee charged by the fund provider to the hedge fund."],
             ["Where the Bank acts as Investment Advisor, it receives an Investment Advisor Fee from the fund house or investment manager in addition to trailer fees, ranging from 0.1% to 0.5% per annum of the fund’s Assets Under Management."]])]],

        ["What is an Individual Client Segregated Account at Standard Chartered?",
         [["Under the Central Securities Depositories Regulation, the Bank, in its capacity as custodian and direct participant, must offer clients the choice between an Omnibus Client Segregated Account and an Individual Client Segregated Account (ICSA) at each Central Securities Depository within the European Economic Area, for Mutual Funds, Hedge Funds, Structured Notes and Fixed Income products. The Bank may not be able to offer ICSA for all products, and ICSA is subject to fees."]]],

        ["What fees apply to a Standard Chartered ICSA for Fixed Income and Structured Products?",
         ["The fees are:",
          B([["Account set up: USD 82 one-time per account."],
             ["Portfolio and transaction fees: aggregate fees at a minimum of USD800 per month, comprising a portfolio fee ranging from 1bps to 6bps per annum depending on the markets of the securities, applied on the market value of the securities on the last business day of the month; and settlement and transaction handling ranging from USD6 to USD200 depending on the markets of the securities."],
             ["Processing and operational fees: USD280 per month."],
             ["Third-party expenses: ranging from USD5 to USD100 per corporate action, including but not limited to sub-custodian fees, depository charges, security pricing charges, stamp duties, mailing, communication and printing costs."]])]],

        ["What fees apply to a Standard Chartered ICSA for Mutual Funds and Hedge Funds?",
         ["The fees are:",
          B([["Account set up: USD 82 one-time per account. Each dividend option requires an individual account."],
             ["Account maintenance: USD 1,290 per month."],
             ["Transaction fee, on top of any upfront fee charged: USD equivalent of EUR 24 per transaction."],
             ["Corporate action fee: USD equivalent of EUR 24 per corporate action."],
             ["Third-party charges and any out-of-pocket expenses paid by the Bank may apply in addition."]])]],
      ]},

      { name: "Insurance, Wealth Lending and Proceeds", url: PG, qas: [
        ["What fees does Standard Chartered receive on insurance products?",
         ["Where the Bank acts as referrer or distributor to the client, it receives the following from a third party or product provider:",
          B([["Life Insurance: up to the Total Distribution Cost To Date as set out in the Policy Illustration section of the insurance plan."],
             ["General Insurance and Credit Life Insurance: from 10% to 40% of the premium."]])]],

        ["What interest applies to a Standard Chartered Wealth Lending Overdraft Facility?",
         [["The Wealth Lending Overdraft Facility, available in USD, SGD, HKD, JPY, EUR, GBP, CHF, AUD, CNH or NZD, carries up to a 1.90% spread per annum plus the base rate of the respective currency’s loan amount, subject to the respective account’s minimum unarranged overdraft interest amount as stated in the Pricing Guide. The base rate refers to the applicable benchmark lending rate or an internally set lending rate."]]],

        ["What fees apply when Standard Chartered investment proceeds are paid by cashier’s order or bank draft?",
         ["These fees are deducted directly from dividends or redemption proceeds, and the net amount is returned by cashier’s order or bank draft. They can be avoided by opening a current or savings account in the respective currency for crediting:",
          B([["Cashier’s Order, in S$ only: S$40.00 per Cashier’s Order."],
             ["Bank Draft, in foreign currency only: 1/8% commission, subject to a minimum of US$40.00 and a maximum of US$70.00 per Bank Draft, plus 1/8% in lieu of exchange, subject to a minimum of US$10.00 and a maximum of US$700.00."]])]],
      ]},
    ],
  },

  // =============================================== REMITTANCES AND PAYMENTS
  {
    name: "Remittances and Payments",
    subs: [
      { name: "Telegraphic Transfers", url: PG, qas: [
        ["What fees apply to Standard Chartered Inward Telegraphic Transfers?",
         ["The fees for Inward Telegraphic Transfers are:",
          B([["Credit to an SGD account, in SGD: S$10. Credit to an SGD account, in foreign currency: free."],
             ["Credit to a foreign currency account, in SGD: free."],
             ["Credit to a foreign currency account, in foreign currency to a same currency account: S$10 or its equivalent."],
             ["Credit to a foreign currency account, in foreign currency to a different currency account: free."]])]],

        ["What fees apply to a Standard Chartered Outward Telegraphic Transfer debiting an SGD account?",
         ["The Personal Banking fees are:",
          B([["Handling commission, online: 1/16% commission, minimum S$15, maximum S$50. Non-online: 1/8% commission, minimum S$30, maximum S$100."],
             ["Reimbursement cover: S$20 for an SGD transfer and US$15 for a foreign currency transfer, online or non-online."],
             ["Overseas bank charges: S$50 online, or as applicable non-online."],
             ["Via the MAS Electronic Payment System (MEPS+): free via FAST online, or S$20 non-online."],
             ["For Priority Banking, handling commission is waived online and charged at 1/16% commission, minimum S$15, maximum S$50, non-online."]])]],

        ["What fees apply to a Standard Chartered Outward Telegraphic Transfer debiting a foreign currency account?",
         ["The Personal Banking fees, which apply to foreign currency accounts including Time Deposits, are:",
          B([["Handling commission online: 1/16% commission, minimum US$10, maximum US$35, plus 1/16% in lieu of exchange, minimum US$5, maximum US$100."],
             ["Handling commission non-online: 1/8% commission, minimum US$20, maximum US$70, plus 1/8% in lieu of exchange, minimum US$10, maximum US$100."],
             ["Reimbursement cover: S$20 for an SGD transfer and US$15 for a foreign currency transfer."],
             ["Overseas bank charges: S$50 online, or as applicable non-online."],
             ["For Priority Banking, handling commission is waived online."]])]],

        ["What waivers apply to Priority Banking clients on Standard Chartered Outward Telegraphic Transfers?",
         ["Handling commission and commission in lieu are waived for Priority Banking customers on online Outward Telegraphic Transfers. For Priority customers, fund transfers between Standard Chartered accounts include waivers of handling commission, commission in lieu and reimbursement cover, though overseas bank charges remain applicable if any. FAST allows a maximum of up to S$200,000 per day."]],

        ["Who bears the charges on a Standard Chartered online telegraphic transfer?",
         ["For online transfers, the remitter may choose between two options:",
          B([["All charges, meaning local bank charges and overseas bank charges, are borne by the remitter. The remittance amount and all charges are debited from the Standard Chartered account specified as the source of funds."],
             ["All charges are borne by the payee as specified by the remitter. Only the remittance amount is debited, all charges are deducted from the remittance amount, and the payee receives the balance after deduction."],
             ["Local bank charges include handling commissions, commissions in lieu and reimbursement cover. Overseas bank charges include correspondent bank and beneficiary bank charges, if any."],
             ["Reimbursement cover is applicable regardless of the destination of the transfer, and fees imposed by other banks may apply and are borne by the remitter."]])]],
      ]},

      { name: "Drafts and Cheques", url: PG, qas: [
        ["What fees apply to inward drafts and cheques at Standard Chartered?",
         ["The fees are:",
          B([["Correspondent bank drafts drawn on Standard Chartered Bank (Singapore) Limited, credited in SGD or foreign currency: free."],
             ["SGD bank drafts drawn by an overseas banking affiliate of Standard Chartered on Standard Chartered Bank (Singapore) Limited: free."],
             ["Foreign currency drafts or cheques drawn on banks in Singapore: free for USD, and S$10 for other currencies."]])]],

        ["What fees apply to foreign currency drafts or cheques drawn on overseas banks at Standard Chartered?",
         ["The fees, which are the same for Personal and Priority Banking, are:",
          B([["Credit to an SGD account: 1/8% commission, minimum S$25, maximum S$200, including postage."],
             ["Credit to a foreign currency account: 1/8% commission, minimum US$15, maximum US$120, including postage."],
             ["Cheque return fee: S$40."],
             ["For foreign currency drafts and cheques drawn on overseas banks and banks not participating in the Cheque Truncating System, drawee bank charges and collection bank charges are also applicable if any."]])]],

        ["What fees apply to issuing an outward draft at Standard Chartered?",
         ["The issuance fees, which are the same for Personal and Priority Banking, are:",
          B([["Debit an SGD account: 1/8% commission, minimum S$30, maximum S$100."],
             ["Debit another foreign currency account, including Time Deposits: 1/8% commission, minimum US$20, maximum US$70, plus 1/8% in lieu of exchange, minimum US$10, maximum US$700."]])]],

        ["What fees apply to amending, cancelling or investigating a Standard Chartered draft?",
         ["The fees, which are the same for Personal and Priority Banking, are:",
          B([["Cable cost and tracers for draft investigations: S$20 per SGD draft, or US$15 per foreign currency draft."],
             ["Amend or replace a draft: S$30 for an SGD draft or US$20 for a foreign currency draft, plus cable cost if applicable."],
             ["Cancel or refund a draft: S$20 credited to an SGD account or US$15 credited to a foreign currency account, plus cable cost if applicable."]])]],
      ]},

      { name: "Cheque Charges", url: PG, qas: [
        ["What cheque clearing fee does Standard Chartered charge?",
         ["The cheque clearing fee is S$3 per SGD cheque and US$3 per USD cheque. In addition:",
          B([["Effective 1 November 2023, the Bank charges payers, meaning individuals who issued a cheque, a cheque clearing fee for all cheques processed."],
             ["There are no fees for Retail Banking clients who are payees, meaning individuals receiving a cheque as payment, until further notice."],
             ["Cheque clearing fees are waived for individual clients who are the main or primary account holder and who are 60 years old as of 31 December 2025."]])]],

        ["What fees apply to stopping a payment or a returned cheque at Standard Chartered?",
         ["The fees are S$40 per SGD cheque and US$30 per foreign currency cheque, applying equally to stop payment and to returned cheques."]],

        ["What does Standard Chartered charge for cheque image retrieval?",
         ["The cheque image retrieval fee depends on the clearing date:",
          B([["Clearing date within 1 year: S$20."],
             ["Between 1 and 3 years: S$30."],
             ["More than 3 years: S$50."]])]],

        ["What does Standard Chartered charge for direct marking and immediate clearing of cheques?",
         ["The fees are:",
          B([["Direct marking of SGD cheques: S$100."],
             ["Immediate clearing of in-house cheques over the counter: S$30 per SGD cheque, or US$20 per foreign currency cheque."]])]],
      ]},

      { name: "Cashier’s Orders, FAST, GIRO and Standing Instructions", url: PG, qas: [
        ["What does Standard Chartered charge for issuing a Cashier’s Order?",
         ["Cashier’s Orders are in SGD only. The issuance fees are:",
          B([["Debiting an SGD account, payment to own name: the first Cashier’s Order is free, thereafter S$5 per Cashier’s Order per transaction."],
             ["Debiting an SGD account, payment to third party names: S$5. Replacement or amendment: S$10."],
             ["Debiting a foreign currency account, payment to own name: the first Cashier’s Order is free, thereafter US$3 per Cashier’s Order per transaction."],
             ["Debiting a foreign currency account, payment to third party names: US$3 or its equivalent. Replacement or amendment: US$10."],
             ["Cancel or refund: S$5 crediting an SGD account, or US$3 crediting a foreign currency account."]])]],

        ["What does Standard Chartered charge for FAST and GIRO?",
         ["The fees are:",
          B([["FAST incoming transfer: free."],
             ["GIRO arrangement: free."],
             ["GIRO return: S$10 per return."],
             ["GIRO amendment: S$10 per amendment."]])]],

        ["What does Standard Chartered charge for Standing Instructions?",
         ["The fees, which are the same for Personal and Priority Banking, are:",
          B([["Automated setup: free online, or S$10 per setup non-online."],
             ["Amendment: free online, or S$10 per amendment non-online."],
             ["Set up for an Outward Telegraphic Transfer: S$50 per setup."],
             ["Manual monitoring: S$50 per payment."],
             ["Standing Instruction return due to insufficient funds: S$50."]])]],
      ]},
    ],
  },

  // ======================================================= CARDS AND LOANS
  {
    name: "Cards and Loans",
    subs: [
      { name: "Credit Cards — Interest and Charges", url: PG, qas: [
        ["What late payment charge applies to a Standard Chartered credit card?",
         [["A late payment charge of S$100 is charged if the minimum payment due, as indicated on the statement, is not received by the Bank by the due date. The same S$100 charge applies to Credit Card Funds Transfer accounts."]]],

        ["What cash advance fee applies to a Standard Chartered credit card?",
         [["The cash advance fee per transaction is 8% on the cash advance, subject to a minimum fee of S$15, in addition to finance charges of 0.082% per day on the amount withdrawn from the date of the transaction until the date of full payment, giving an Effective Interest Rate of 29.9% per annum. With effect from 15 July 2023, quasi-cash transactions under MCC category 7995 are regarded as cash advances."]]],

        ["What finance charges apply to purchases on a Standard Chartered credit card?",
         [["The Effective Interest Rate is 27.9% per annum minimum. If payment is not made in full by the due date, finance charges are billed on the next statement. Finance charges are calculated on a daily basis at 0.076% on the outstanding balance, including purchase transactions and finance charges resulting from those transactions, from the date of the unpaid transactions until full payment is received, and on all new transactions from their respective transaction dates. No finance charge is levied if payment is received in full by the due date and there is no balance carried forward from the previous statement."]]],

        ["What finance charges apply to the Standard Chartered Platinum Access Card and Automatic Flexible Payment Scheme?",
         ["The charges are:",
          B([["Effective Interest Rate of 27.9% per annum minimum, with finance charges calculated daily at 0.076% on the outstanding balance if payment is not made in full by the due date."],
             ["Effective interest rate for instalments under the Scheme: 9.32% per annum. If the instalment amount payment is not received in full on or before the due date and a balance is carried forward, finance charges are calculated daily at the minimum Effective Interest Rate of 27.9% per annum, which is 0.076% per day."],
             ["An administrative fee of 6% is charged on all approved credit card transactions under the Scheme."]])]],

        ["What transfer fees apply to Standard Chartered rewards points?",
         ["The transfer fees are:",
          B([["KrisFlyer Miles: S$27.25 including 9% GST."],
             ["SC EasyRewards (Transfer Rewards): S$27.25 including 9% GST."],
             ["360° Rewards Points transfer: S$10 for every 100,000 points transferred to each nominee’s credit card account, not applicable for cardholders of the Visa Infinite Credit Card. A nominee is a nominated family member or friend holding a credit card eligible to earn 360° Rewards Points, where that card is valid, subsisting and in good standing."]])]],

        ["What early redemption fee applies to Standard Chartered EasyPay on Retail Balances?",
         [["The early redemption fee is 3% of the unbilled outstanding principal."]]],
      ]},

      { name: "Credit Cards — Annual Membership Fees", url: PG, qas: [
        ["What annual fees apply to Standard Chartered’s premium credit cards?",
         ["The annual membership fees are:",
          B([["Beyond Credit Card: S$1,635 including 9% GST. This annual fee is strictly non-waivable. Supplementary cards are free, up to 4 cards."],
             ["Visa Infinite Credit Card: S$599.50 including 9% GST. This annual fee is strictly non-waivable. Supplementary cards are free, up to 4 cards."],
             ["Priority Banking Visa Infinite Credit Card: S$327 including 9% GST, strictly non-waivable, though waived for the first year. Supplementary card: S$163.50 including 9% GST, free for the first year, up to 4 cards."],
             ["Preferred World: S$218 including 9% GST. This annual fee is strictly non-waivable. Supplementary cards are free, up to 4 cards."]])]],

        ["What annual fees apply to Standard Chartered’s mainstream credit cards?",
         ["The annual membership fees, all with free supplementary cards up to 4 cards, are:",
          B([["Smart Credit Card: S$99.19 including 9% GST from 6 December 2025, free for the first year."],
             ["Rewards+, Simply Cash, Spree, Prudential Platinum, Prudential Visa Signature and NUS Alumni Platinum Credit Cards: S$196.20 including 9% GST, free for the first year."],
             ["Journey Credit Card: S$196.20 including 9% GST, free for the first year if FEE WAIVED is selected."],
             ["MANHATTAN Platinum and MANHATTAN World Mastercard: free for the first year."],
             ["Business Platinum Credit Card: S$163.50 including 9% GST, free for the first three years."]])]],

        ["What annual fees apply to the Standard Chartered Platinum, Gold, Classic and S$500 cards?",
         ["The annual membership fees, all with free supplementary cards up to 4 cards, are:",
          B([["Platinum Visa/Mastercard Credit Card: free for the first three years. Where the Platinum Visa Credit Card is issued upon approval and disbursement of a CashOne Personal Loan, the annual fee is free for 5 years. Where the Platinum Mastercard Credit Card is issued upon successful application of a Debt Consolidation Plan, the annual fee is free."],
             ["Gold Card: S$190.75 including 9% GST, strictly non-waivable."],
             ["Classic Card: S$65.40 including 9% GST, strictly non-waivable."],
             ["S$500 Cards: S$32.70 including 9% GST, free for the first year."],
             ["Annual fee is applicable for both active and inactive credit cards."]])]],
      ]},

      { name: "Credit Cards — Other Fees", url: PG, qas: [
        ["What other fees and charges apply to a Standard Chartered credit card?",
         ["The other fees are:",
          B([["Overlimit fee: S$40, charged to the card account if the current balance exceeds the credit limit."],
             ["Sales draft: S$5."],
             ["Rejected Direct Debit Authorisation or GIRO payment due to insufficient funds: S$50."],
             ["Returned cheque due to insufficient funds: S$40."],
             ["Branch counter payment: S$5.45 per transaction including GST."],
             ["Request for a past copy of a statement, both paper and eStatement: S$20 per monthly statement for current to 12 months, S$50 per monthly statement beyond 12 months, capped at S$500 per request."]])]],

        ["When is the Standard Chartered branch counter payment fee not charged?",
         [["The S$5.45 branch counter payment fee applies where the instruction to make a payment to a credit card account is submitted at any of the Bank’s branches, except where the account holder is at least 65 years old at the time the payment instruction is submitted at a branch."]]],
      ]},

      { name: "Credit Cards — Important Notes", url: PG, qas: [
        ["What is the repayment period on a Standard Chartered credit card?",
         [["The repayment period is at least 22 days from the billing or statement date."]]],

        ["How are Standard Chartered credit card finance charges calculated?",
         [["Finance charges are imposed on the outstanding balance on the statement if payment is not made in full by the due date, calculated from the date of the transaction to the date of full payment. From 15 July 2022, no finance charges are imposed on any fees billed to the card account, such as membership fee, late fee and any other fees charged by the Bank."]]],

        ["What is the minimum payment due on a Standard Chartered credit card?",
         ["The minimum payment due applicable to all credit card accounts, including Credit Card Funds Transfer accounts, is:",
          B([["The greater of either S$50 or 1% of principal, including any instalments billed in the current month; plus"],
             ["Interest, fees and charges; and"],
             ["Overlimit amount and past due amount, if any."],
             ["Outstanding balances below S$50 must be repaid in full."]])]],

        ["What is a Standard Chartered cardholder’s liability for a lost or stolen credit card?",
         [["If the credit card is lost or stolen or if the PIN is disclosed, the cardholder’s maximum liability is capped at S$100, provided the cardholder immediately notifies the Bank in writing, assists in the recovery, and furnishes a statutory declaration or police report together with any other information required, and the Bank is satisfied that the loss, theft or disclosure of the card or PIN is not due to the cardholder’s negligence or default. This limitation of liability does not apply to cash advances."]]],

        ["How are foreign currency transactions converted on a Standard Chartered card?",
         [["All foreign currency transactions effected in US Dollars are converted to Singapore Dollars on the date of conversion. Transactions effected in a currency other than US Dollars are converted into US Dollars before being converted into Singapore Dollars. The exchange rate may differ from the rate in effect on the transaction date due to market fluctuations. Any rate imposed is final and conclusive, and the cardholder bears all exchange risks, loss, commission and other bank costs. All conversions are based on prevailing wholesale interbank rates or the government-mandated rate, as determined by Visa or Mastercard."]]],

        ["What charges apply to foreign currency transactions on a Standard Chartered card?",
         ["All foreign currency transactions, including overseas and online transactions, charged to Visa or Mastercard cards are subject to:",
          B([["A prevailing charge of 1% of the converted Singapore Dollar amount, representing the charge imposed by Visa or Mastercard on the transaction."],
             ["A prevailing fee of 2.5% of the converted Singapore Dollar amount."]])]],

        ["What is Dynamic Currency Conversion and what does it cost on a Standard Chartered card?",
         [["Dynamic Currency Conversion (DCC) is a service offered at selected overseas ATMs, websites or by certain merchants, which converts foreign currency transactions into Singapore Dollars. Where DCC is used, the process of conversion and the exchange rate applied are determined by the relevant DCC service provider and not by Standard Chartered. All transactions converted via DCC are subject to a prevailing charge of 1% of the transaction amount, representing the charge imposed by Visa or Mastercard."]]],

        ["Are Singapore Dollar transactions processed overseas charged on a Standard Chartered card?",
         [["Yes. Where a client enters into Singapore Dollar transactions with a local merchant that routes its payment processing through an overseas intermediary, or with a merchant registered by its acquiring bank as having been acquired overseas regardless of where the merchant is actually located, such transactions are subject to a prevailing charge of 1% of the transaction amount, representing the charge imposed by Visa or Mastercard. In both cases the transaction is treated as a foreign currency transaction, and the conversion process and exchange rate applied are determined by the payment processing intermediary or the Acquirer, not by the Bank."]]],
      ]},

      { name: "Personal Loans", url: PG, qas: [
        ["What fees apply to a Standard Chartered CashOne, Credit Card Instalment Loan, Dash Advance Personal Loan or ezyCash?",
         ["The fees are:",
          B([["Annual fee, first year: S$199, deducted upfront from the approved CashOne loan amount. Second year onwards, until either the expiry of the instalment tenure or the final instalment payment, whichever is earlier: S$50 charged annually, but only if any minimum payment due amounts were not paid on or before the due dates for the preceding 12 consecutive calendar months."],
             ["Late payment charge: S$100 if the minimum payment due amount is not received by the due date."],
             ["Finance charges: if the instalment amount payment is not received in full on or before the due date and a balance is carried forward, finance charges are calculated daily at the minimum Effective Interest Rate of 29.9% per annum, which is 0.082% per day."],
             ["Early redemption fee: S$150 or 3% of the outstanding principal, whichever is higher."],
             ["Change of tenure: S$50."],
             ["Default interest: if the minimum payment due amount is not received on or before the due date twice within any consecutive 6-month period, a default interest of 4% per annum is added to the original Effective Interest Rate on the entire outstanding balance."],
             ["Branch counter payment: S$5.45 per transaction inclusive of GST, except where the account holder is at least 65 years old at the time the payment instruction is submitted at a branch."]])]],

        ["What is the minimum payment due on a Standard Chartered Personal Loan account?",
         ["The minimum payment due applicable to all Personal Loan accounts is:",
          B([["The approved monthly instalment amount; plus"],
             ["Interest, fees and charges; plus"],
             ["Overlimit and past due amount, if any."],
             ["Outstanding balances below S$50 must be repaid in full."],
             ["From 15 June 2023, only 1% of the approved monthly principal instalment amount, excluding interest, fees and charges, is included in the minimum payment due for selected accounts. The Bank reserves the right to determine at its sole and absolute discretion whether to offer this revised calculation."]])]],

        ["What fees apply to the Standard Chartered Interest Free CashOne Personal Loan?",
         ["The fees are:",
          B([["Annual fee: 4.5% of the approved loan amount, deducted upfront from the approved loan amount."],
             ["Late payment charge: S$100 if the minimum payment is not received by the due date."],
             ["Finance charges: if the instalment amount payment is not received in full on or before the due date and a balance is carried forward, finance charges are calculated daily at the minimum Effective Interest Rate of 29.9% per annum, which is 0.082% per day."],
             ["Default interest: if the instalment amount payment is not received in full on or before the due date twice within any consecutive 6-month period, a default interest of 4% per annum is added to the original Effective Interest Rate on the entire outstanding balance."],
             ["Branch counter payment: S$5.45 for each payment made over the branch counter, inclusive of GST."]])]],

        ["What is the minimum payment due on a Standard Chartered Interest Free CashOne Personal Loan?",
         ["The minimum payment due is:",
          B([["The statement-billed monthly instalment amount; or"],
             ["1% of the approved monthly principal instalment amount, or S$50, whichever is greater;"],
             ["Plus any interest, fees, charges and any overlimit and past due balance."],
             ["Outstanding balances below S$50 must be repaid in full."]])]],

        ["What fees apply to the Standard Chartered Debt Consolidation Facility?",
         ["The fees are:",
          B([["Joining fee: S$199, a one-time charge."],
             ["Late payment charge: S$100 if the minimum payment is not received by the due date."],
             ["Finance charges: if the full instalment payment amount is not received by the due date, the outstanding balance is carried forward to the next statement and finance charges are calculated daily at the minimum Effective Interest Rate of 26.9% per annum, which is 0.074% per day."],
             ["Early redemption fee: S$250 or 5% of the outstanding principal, whichever is higher."],
             ["Branch counter payment: S$5.45 per transaction inclusive of GST, except where the account holder is at least 65 years old at the time the payment instruction is submitted at a branch."]])]],

        ["What is the minimum payment due on a Standard Chartered Debt Consolidation Facility?",
         ["The minimum payment due applicable to all Debt Consolidation Facility accounts is:",
          B([["The approved monthly instalment amount; plus"],
             ["Interest, fees and charges; plus"],
             ["Overlimit and past due amount, if any."],
             ["Outstanding balances below S$50 must be repaid in full."]])]],
      ]},

      { name: "Personal Credit and Lines of Credit", url: PG, qas: [
        ["What fees and interest apply to Standard Chartered Personal Credit, Salary Advance and Preferred Line of Credit?",
         ["The fees and interest are:",
          B([["Annual fee: S$80."],
             ["Prevailing interest rate: 18.88% per annum for Personal Credit, Credit Manhattan and Credit One. For Salary Advance and Preferred Line of Credit, the prevailing interest rate is 6.88% per annum or 17.9% per annum."],
             ["Late payment charges: S$80 per month if the payment due date is missed, or if less than the minimum repayment amount is paid."],
             ["Overlimit charges: 5% per annum in addition to the prevailing interest rate on the over limit balances."]])]],

        ["What cash withdrawal fees apply to a Standard Chartered Personal Credit or line of credit account?",
         ["The withdrawal fees are:",
          B([["ATM cash withdrawal: free at Standard Chartered ATMs and on the atm5 network. atm5 is the shared ATM network giving island-wide access to ATMs bearing the atm5 logo from Standard Chartered, Bank of China, Citibank, HSBC, Maybank and State Bank of India."],
             ["Overseas cash withdrawal: 2% of the amount withdrawn, subject to a minimum of S$5 and maximum of S$60 per withdrawal, or a minimum of US$5 and maximum of US$60 per withdrawal for USD accounts."]])]],

        ["What other fees apply to a Standard Chartered Personal Credit or line of credit account?",
         ["The other fees are:",
          B([["Returned cheque fee: S$40 per cheque."],
             ["Stop cheque fee: S$40 per cheque."],
             ["GIRO returned fee: S$10 per transaction."],
             ["Lost card replacement fee: S$5 per card."]])]],

        ["What is the maximum credit limit on a Standard Chartered Personal Credit account?",
         [["The maximum credit limit is 2 times the client’s current monthly salary, up to S$100,000, whichever is lower. The Bank reserves the right to grant or assign a lower credit limit."]]],

        ["What is the minimum monthly repayment on a Standard Chartered Personal Credit account?",
         ["If the outstanding balance is S$50 or more, the minimum monthly repayment is the higher of:",
          B([["S$50; or"],
             ["1% of the outstanding principal amount plus interest, fees and charges, plus any amount in the account balance exceeding the credit limit, and any past due amount."],
             ["If the outstanding balance is less than S$50, the entire outstanding balance is payable by the due date."]])]],

        ["Is interest paid on a credit balance in a Standard Chartered Personal Credit account?",
         [["Where there is a credit balance in the account, the client may be entitled to receive interest depending on the type of account. The rate of interest may be fixed or varied as the Bank determines, and the Bank pays interest monthly or at other regular intervals that it determines."]]],
      ]},

      { name: "Mortgage Loans", url: PG, qas: [
        ["Does the Standard Chartered Pricing Guide cover mortgage interest rates?",
         [["No. The Pricing Guide does not cover interest rates on Mortgage Facilities. Applicable interest rates, fees and other conditions are set out in the borrower’s Facility Letter, and the terms in the Facility Letter apply if there is any inconsistency with the Pricing Guide."]]],

        ["What redemption and cancellation fees apply to a Standard Chartered mortgage loan?",
         ["The fees are:",
          B([["Partial redemption fee: 1.5% based on the amount of Mortgage Loan redeemed during the Lock-In Period. A loan may be redeemed partially by serving 1 month’s notice in writing or by paying 1 month’s interest in lieu of notice, and the minimum partial redemption amount is S$10,000."],
             ["Full redemption fee: 1.5% based on the amount of Mortgage Loan redeemed during the Lock-In Period. A loan may be redeemed in full by serving 2 months’ notice in writing or by paying 2 months’ interest in lieu of notice."],
             ["Cancellation fee: 1.5% based on the amount of Mortgage Loan cancelled or deemed cancelled by the borrower."]])]],

        ["What other fees apply to a Standard Chartered mortgage loan?",
         ["The other fees are:",
          B([["Repricing fee: an administrative fee of S$1,000."],
             ["Late payment fee: S$50."],
             ["Default rate: 5% above the SGD Prime lending rate of 5.75% on the overdue instalment amount."],
             ["Request for documents such as a Facility Letter or Supplemental Facility Letter: S$54.50 per copy inclusive of GST."],
             ["Request for a past copy of statements: S$20 per monthly statement for current to 12 months, S$50 per monthly statement beyond 12 months, capped at S$500 per request."]])]],
      ]},

      { name: "Auto Financing", url: PG, qas: [
        ["What fees apply to Standard Chartered Auto Financing?",
         ["The fees are:",
          B([["Late payment fee: S$60."],
             ["Late payment interest: 12% per annum on the overdue amount, billed on a monthly basis."],
             ["Photocopy of hire purchase agreement or log card: S$54.50 per copy inclusive of GST, waived for the first year from the commencement date."],
             ["Request for a past copy of statements: S$20 per monthly statement for current to 12 months, S$50 per monthly statement beyond 12 months, capped at S$500 per request."]])]],

        ["What early completion fee applies to Standard Chartered Auto Financing?",
         ["The early completion fee depends on how long the agreement has run:",
          B([["Within 12 months from the commencement date: 20% of outstanding interest plus 1.5% of balance payable."],
             ["12 months or more from the commencement date: 20% of outstanding interest plus 1% of balance payable."]])]],

        ["How is an early full settlement of Standard Chartered Auto Financing requested?",
         ["An early full settlement or settlement quote is requested as follows:",
          B([["Customers submit a signed and completed Early Full Settlement form. Car dealers submit an Early Full Settlement form and a Dealer’s Indemnity form."],
             ["The completed form is mailed or dropped off at any local Standard Chartered branch."],
             ["Requests for early full settlement quotes are processed and mailed within 2 business days."],
             ["For immediate payment of the full settlement amount, a FAST transfer can be made directly to the repayment account. Payment can also be dropped off via Cashier’s Order, cheque or cash at any local branch, with a processing time of 3 business days."],
             ["Instructions are processed upon receipt of the full payment and the completed settlement form."]])]],
      ]},

      { name: "ATM and Debit Card", url: PG, qas: [
        ["What overseas transaction fee applies to a Standard Chartered debit card?",
         [["The debit card overseas transaction fee is up to 3.5% of the amount transacted via Mastercard, for both Personal Banking and Priority Banking."]]],

        ["What does Standard Chartered charge for cash withdrawal at Cirrus ATMs worldwide?",
         ["For Personal Banking, the fee is 2% of the amount withdrawn, subject to a minimum of S$5 and maximum of S$60, or a minimum of US$5 and maximum of US$60 for USD accounts. It is waived for Priority Banking. In addition:",
          B([["All such overseas cash withdrawals are also subject to additional fees imposed by the other bank and by Visa or Mastercard International."],
             ["For credit cards linked to a current or savings account, the withdrawal amount is deducted from the credit card if the overseas ATM does not support account type selection, and the cash advance fee applies."],
             ["The cash advance fee per transaction is 8% on the cash advance, subject to a minimum fee of S$15, in addition to finance charges of 0.082% per day on the amount withdrawn from the transaction date until full payment, giving an Effective Interest Rate of 29.9% per annum."],
             ["World Partner customers enjoy the first cash withdrawal per month free of charge at Standard Chartered and Cirrus ATMs worldwide when using their World Partner Platinum Debit Card. This waiver does not include fees levied by other banks in Singapore or overseas for accessing their ATM networks."]])]],
      ]},
    ],
  },

  // ============================================== BRANCH AND OTHER SERVICES
  {
    name: "Branch and Other Services",
    subs: [
      { name: "Coin and Notes Services", url: PG, qas: [
        ["What does Standard Chartered charge for coin and notes services?",
         ["The Personal Banking fees, all of which are waived for Priority Banking, are:",
          B([["Deposit or withdrawal of coins: S$5 for between S$20 and S$100, and S$10 for every S$100 or part thereof."],
             ["Changing of notes to coins: S$5 per S$50 or part thereof."],
             ["Notes deposit: free for the first 300 pieces, then S$2 for every subsequent 100 pieces or part thereof, in all denominations."]])]],

        ["What conditions apply to Standard Chartered coin and notes services?",
         ["The conditions are:",
          B([["Charges are on a per account, per day basis."],
             ["Charges are waived for e$aver Kids account transactions."],
             ["Coin services are only available on Tuesdays and Thursdays, from 11am to 1pm at selected branches located in shopping malls, and from 10am to 12pm at all other branches."]])]],
      ]},

      { name: "Safe Deposit Box", url: PG, qas: [
        ["What fees apply to a Standard Chartered safe deposit box?",
         ["The fees are:",
          B([["Key deposit fee, refundable: S$150."],
             ["Addition, change or revocation of nomination: S$10 each."],
             ["Change of safe deposit box: S$20."],
             ["Break box charges due to a lost key, where the key deposit fee becomes non-refundable: S$210."],
             ["Inventory for estate duty or deceased customers: S$30."]])]],

        ["What annual rental applies to a Standard Chartered safe deposit locker?",
         ["Annual rental by locker size, excluding GST, is as follows for Personal Banking and Priority Banking respectively:",
          B([["4.75″ x 3.5″ x 20.0″: S$150 and S$75. 7.5″ x 3.5″ x 20.0″: S$250 and S$125."],
             ["8.0″ x 4.0″ x 24.0″: S$300 and S$150. 10.0″ x 5.0″ x 24.0″: S$450 and S$225."],
             ["7.5″ x 7.5″ x 18.75″: S$400 and S$200. 15.0″ x 7.5″ x 18.75″: S$600 and S$300."],
             ["16.0″ x 8.0″ x 24.0″: S$700 and S$350. 20.0″ x 10.0″ x 24.0″: S$800 and S$400."]])]],

        ["What conditions apply to Standard Chartered safe deposit locker rental?",
         ["The conditions are:",
          B([["The annual rental fee is subject to the prevailing Goods and Services Tax."],
             ["With effect from 1 January 2024, the GST rate was adjusted from 8% to 9%. Any outstanding rental fees not paid by 31 December 2023 are subject to 9% GST from 1 January 2024 onwards."],
             ["The Safe Deposit Locker service is only available at the Battery Road Branch."]])]],
      ]},

      { name: "Guarantees", url: PG, qas: [
        ["What does Standard Chartered charge for a banker’s guarantee?",
         ["The fees, which are the same for Personal and Priority Banking, are:",
          B([["Banker’s guarantee: 1% per annum or a minimum of S$100, whichever is higher, subject to a minimum loan limit of US$20,000 or equivalent for a maximum of 1 year, inclusive of the claim period."],
             ["Amendment: S$25 per amendment."],
             ["Cancellation: S$100 per guarantee."]])]],
      ]},

      { name: "Other Services", url: PG, qas: [
        ["What does Standard Chartered charge for cash deposits and withdrawals?",
         ["The fees, which are the same for Personal and Priority Banking where charged, are:",
          B([["Cash deposit in SGD to any account: free."],
             ["Cash deposit in foreign currency to a same currency account: 1.5% repatriation fee, minimum US$15 or its equivalent."],
             ["Cash withdrawal in SGD from any account: free."],
             ["Cash withdrawal in foreign currency from a same currency account: 1.5% commission, minimum US$5 or its equivalent."]])]],

        ["What does Standard Chartered charge for statements and confirmations?",
         ["The fees are:",
          B([["Request for a past copy of a statement, applying to both paper and eStatements: S$20 per monthly statement for current to 12 months, S$50 per monthly statement beyond 12 months, capped at S$500 per request. This applies to both Personal and Priority Banking."],
             ["Deposits and withdrawal confirmation: S$10 per monthly statement for current to 12 months and S$50 beyond 12 months, waived for Priority Banking."],
             ["Statement of interest: S$10 per monthly statement for current to 12 months and S$50 beyond 12 months, waived for Priority Banking."]])]],

        ["What does Standard Chartered charge for certificates, confirmations and reports?",
         ["The fees are:",
          B([["Certificate of balance: S$20 per request, waived for Priority Banking."],
             ["Audit confirmation: S$52 per request in SGD, or US$30 per request for foreign currency accounts, waived for Priority Banking."],
             ["Letter of reference: S$20 per request."],
             ["Personal Data Report: S$10, waived for Priority Banking."]])]],

        ["What does Standard Chartered charge companies for salary crediting?",
         ["The fees, which are the same for Personal and Priority Banking, are:",
          B([["Salary crediting via Payplus: S$0.50 per item, minimum S$30 or its equivalent."],
             ["Manual salary crediting: S$2 per payee, minimum S$30 or its equivalent."]])]],

        ["What is the Standard Chartered Priority Banking service fee, and when is it waived?",
         ["The Priority Banking service fee is S$50 and is not applicable to Personal Banking. It is automatically waived if the client:",
          B([["Maintains a minimum of S$200,000 in deposits and/or investments; or"],
             ["Maintains a minimum of S$1.5 million in housing loans with the Bank."]])]],

        ["What happens if Standard Chartered Priority Banking eligibility criteria are not met?",
         ["Priority Banking membership is subject to meeting the Bank’s eligibility criteria. Where the criteria are not met, the Bank may at its discretion:",
          B([["Impose a fee, which will be deducted from any of the client’s accounts with the Bank, which the client expressly authorises."],
             ["Convert any or all household members accorded Priority Banking membership through Household Recognition to a Personal Banking relationship."],
             ["Cease or suspend the provision of any Priority Banking services until the eligibility criteria are met."],
             ["Priority Banking Services and Privileges International Terms and Conditions apply, published at ", ["sc.com/sg/terms-and-conditions/priority", L.priorityTnc], "."]])]],
      ]},

      { name: "Important Notes on the Pricing Guide", url: PG, qas: [
        ["What should the Standard Chartered Pricing Guide be read alongside?",
         [["The Pricing Guide should be read in conjunction with the Bank’s Customer Terms, Current/Cheque/Savings Account and Time Deposit Terms, Personal Loan/Personal Line of Credit/Overdraft Terms, Governing Debt Consolidation Facility Terms, Credit Card Terms, and such other terms as may be applicable or otherwise amended from time to time at the Bank’s sole and absolute discretion."]]],

        ["Are the fees in the Standard Chartered Pricing Guide fixed?",
         [["No. All the information in the Pricing Guide is subject to the terms of the banking agreement, and all fees and charges are subject to change. Unless otherwise stated, the meaning of key words and other words used in the banking agreement is explained in the Customer Terms, product terms, and such other terms as may be made available from time to time."]]],

        ["Does the Standard Chartered Pricing Guide constitute an offer?",
         [["No. The Pricing Guide does not constitute an offer to sell, or a solicitation of an offer to buy or sell, any securities or any other financial instrument, to effect any transaction or to provide any service. Not all products and services are available to citizens or residents of all countries."]]],
      ]},
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Pricing Guide FAQ",
  docTitle: "SCB Bank_Pricing Guide FAQ",
  description: "Standard Chartered Singapore Pricing Guide FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
console.log("categories:", CATEGORIES.length, "subcategories:", CATEGORIES.reduce((n, c) => n + c.subs.length, 0));
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Pricing_Guide_FAQ.docx", buf);
  console.log("written", buf.length);
});
