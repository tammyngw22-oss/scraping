const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

const SC = "https://www.sc.com/sg/";
const U = {
  prIndex: SC + "priority/",
  prSignup: SC + "priority/sign-up/",
  prMgm: SC + "priority/mgm-referral-programme/",
  prPass: SC + "priority/priority-pass-faq/",
  fsIndex: SC + "fraud-scam/",
  fsDisable: SC + "fraud-scam/disable-your-digital-services/",
  fsLoan: SC + "fraud-scam/loan-scam/",
  rwIndex: SC + "rewards-programmes/",
  rw360: SC + "rewards-programmes/360-rewards-redemption/",
  grab: SC + "cards/cards-and-cashone/grab/",
  qoo10: SC + "cards/cards-and-cashone/qoo10/",
  pmIndex: SC + "promotions/",
  pmBirthday: SC + "promotions/birthdaytreats/",
  pmCaltex: SC + "promotions/caltex/",
  pmInstant: SC + "promotions/instant/",
  pmMillion: SC + "promotions/million/",
  pmReferTnc: SC + "promotions/refer-a-friend-tnc/",
  pmCatalogue: SC + "promotions/referral-programme/product-catalogue/",
  pmRefSignup: SC + "promotions/referral-signup/",
  pmGoodLife: SC + "promotions/the-good-life-privileges/",
  about: SC + "about-us/",
  business: SC + "business/",
  cdd: SC + "cddreview/",
  esign: SC + "esign/",
  journeyFaq: SC + "important-information/journey-credit-card-faq/",
  allianz: SC + "insurance/allianz-hospital-income-protect/",
  intl: SC + "international-banking/",
  personal: SC + "personal/",
  inner: SC + "scinnercircle/",
  statements: SC + "statements-refresh/",
};

const L = {
  lounges: "https://prioritypass.com/airport-lounges",
  ppLogin: "https://www.prioritypass.com/en/login",
  moneylenders: "https://rom.mlaw.gov.sg/information-for-borrowers/list-of-licensed-moneylenders-in-singapore/",
  visaInfinite: SC + "credit-cards/visa-infinite-card/",
  journey: SC + "credit-cards/journey-credit-card/",
  beyond: SC + "credit-cards/beyond-credit-card/",
  simplyCash: SC + "credit-cards/simply-cash-credit-card/",
  bonussaver: SC + "save/current-accounts/bonussaver/",
  wealthsaver: SC + "save/current-accounts/wealth-saver/",
};

const CATEGORIES = [
  // ====================================================== PRIORITY BANKING
  {
    name: "Priority Banking",
    subs: [
      { name: "Priority Banking Overview", url: U.prIndex, qas: [
        ["What does Standard Chartered Priority Banking offer?",
         ["Priority Banking gives access to a full suite of wealth solutions to help achieve financial goals, along with tailored advice and strategic market insights from the Bank’s experts. It is powered by timely CIO house views and advanced digital tools."]],

        ["What deposit interest rates are available to Standard Chartered Priority Banking clients?",
         ["For a limited time, Priority Banking clients can access the following rates:",
          B([["Up to 5.85% p.a. interest with ", ["Bonus$aver", L.bonussaver], "."],
             ["Up to 4.00% p.a. interest with ", ["Wealth $aver", L.wealthsaver], ", on both SGD and USD deposit balances."],
             ["Up to 4.20% p.a. interest with a Fresh Funds Time Deposit."]])]],
      ]},

      { name: "Priority Banking Sign-Up Rewards", url: U.prSignup, qas: [
        ["What welcome rewards are available for joining Standard Chartered Priority Banking?",
         ["New Priority Banking clients receive over S$60,000 in welcome gifts. Existing Bank customers are not eligible to participate in the New-To-Bank Sign-Up Promotion."]],

        ["What cash rewards apply when joining Standard Chartered Priority Banking?",
         ["Cash rewards of up to S$4,000 are available, based on the amount of Fresh Funds brought in:",
          B([["Funding of S$200,000: reward value S$350."],
             ["Funding of S$1.5 million: reward value S$3,000."],
             ["Funding of S$3 million: reward value S$4,000."]])]],

        ["What Signature CIO Fund rewards apply when joining Standard Chartered Priority Banking?",
         ["Signature CIO Fund rewards of up to S$8,000 in equivalent value are available, based on the amount of Fresh Funds brought in:",
          B([["Funding of S$200,000: reward value S$700."],
             ["Funding of S$1.5 million: reward value S$6,000."],
             ["Funding of S$3 million: reward value S$8,000."]])]],

        ["What additional reward applies to Accredited Investors joining Standard Chartered Priority Banking?",
         ["Customers who bring in Fresh Funds between S$200,000 and below S$1,000,000 are also eligible to receive an additional S$1,000 Cash Reward or Signature CIO Fund Reward, based on their selection, if they qualify and declare, opt in and maintain their status as an Accredited Investor with the Bank. Funding tiers of S$1 million and above require the client to consent to being treated as an Accredited Investor, where applicable and as determined by the Bank in its sole and absolute discretion."]],

        ["What is the Standard Chartered Priority Banking Wealth Cash Reward?",
         ["The Wealth Cash Reward gives up to S$6,000, calculated as S$200 for every S$50,000 purchased in an Eligible Unit Trust offered by the Bank using Fresh Funds in a single day."]],
      ]},

      { name: "Priority Banking Referral Programme", url: U.prMgm, qas: [
        ["What is the Standard Chartered Priority Banking Referral Programme?",
         ["The Priority Banking Referral Programme lets an existing client refer family and friends to start a Priority Banking relationship with Standard Chartered, with rewards accorded to the referrer when the referral is successful."]],

        ["How does a referred individual start a Standard Chartered Priority Banking relationship?",
         ["A referred individual starts a Priority Banking relationship in two steps:",
          B([["Step 1: Choose and open a current or savings account from the available options, which include Bonus$aver, e$aver and USD$aver."],
             ["Step 2: Bring in a minimum of S$200,000 in fresh funds via FAST transfer."]])]],

        ["Which accounts can a referred Standard Chartered Priority Banking client open?",
         ["A referred individual can open any of the following accounts:",
          B([["Bonus$aver, which offers the Bank’s high interest rates and a multi-currency feature for spending, transacting and investing across 14 currencies."],
             ["e$aver, which earns bonus interest on eligible incremental balances, with no lock-in period and instant account opening using MyInfo."],
             ["USD$aver, which earns higher interest rates without locking in funds and saves on foreign currency conversion on USD-denominated transactions."]])]],
      ]},

      { name: "Priority Pass — Application and Membership", url: U.prPass, qas: [
        ["What is the Priority Pass programme?",
         [["Priority Pass is the world’s leading airport lounge access programme, giving access to more than 1,700 airport lounges worldwide. It is complimentary for all Standard Chartered principal cardholders of the Priority Banking Visa Infinite Credit Card, Visa Infinite Credit Card and Journey Credit Card. The full list of participating airport lounges is at ", ["prioritypass.com/airport-lounges", L.lounges], "."]]],

        ["Who can apply for the Priority Pass programme through Standard Chartered?",
         ["Cardholders of the Standard Chartered Priority Banking Visa Infinite Credit Card, Visa Infinite Credit Card and Journey Credit Card can apply for a Priority Pass membership. Only the principal cardholder is eligible to apply."]],

        ["How is a Priority Pass membership applied for through Standard Chartered?",
         ["The principal card must first be activated, then registration is completed using the mobile number registered with the Bank:",
          B([["Step 1: Send an SMS to 77272 in the format for the card held — SCX PP followed by the last 4 digits of the card number for a Journey Credit Card, SCVI PP for a Visa Infinite Credit Card, or PBVI PP for a Priority Banking Visa Infinite Credit Card. For example: SCX PP 1234."],
             ["Step 2: An automated reply acknowledging receipt of the SMS is sent."],
             ["Step 3: If registration is successful, a one-time unique code is sent to the registered mobile number by SMS within 14 working days."],
             ["Step 4: Enter the one-time unique code in the Priority Pass activation link and complete the application."],
             ["Step 5: Download the Priority Pass mobile app and log in with the User ID."],
             ["Step 6: The digital Priority Pass is ready to use. The physical membership card is dispatched within 5 business days of the application being successfully accepted, through normal mail."]])]],

        ["When is the Priority Pass one-time unique code received after registering with Standard Chartered?",
         ["The automated reply is an acknowledgement that the Bank has received the registration for the complimentary Priority Pass membership. If registration is successful, the one-time unique code is received within 14 working days to activate the Priority Pass membership."]],

        ["What is the purpose of the Priority Pass one-time unique code?",
         ["The one-time unique code is tagged to the number of complimentary visits that each applicant is entitled to. The unique code is strictly valid for one-time usage only."]],

        ["Why might a Standard Chartered Priority Pass registration be rejected, and what are the next steps?",
         ["A rejection SMS states the reason. The next steps depend on the reason:",
          B([["A unique one-time code was issued previously: use the code already provided to register for a Priority Pass membership. If an active Priority Pass membership is already held, a new code will not be issued; log-in details can be reset on the Priority Pass website under the “Forgot Your Details” section."],
             ["Incorrect card details were entered: re-send the SMS registration with the correct last 4 digits of the credit card number, using the mobile number registered with Standard Chartered."],
             ["An incorrect SMS format was entered: re-send the SMS registration using the correct format for the card held."]])]],

        ["Can a Priority Pass membership be applied for with a foreign mobile number?",
         ["No. Only local registered mobile numbers, meaning Singapore phone numbers, can apply for the Priority Pass membership via SMS. The mobile number used to apply must be registered with Standard Chartered before applying for a digital Priority Pass membership."]],

        ["Is a Standard Chartered Priority Pass membership renewed automatically?",
         ["Yes. The Priority Pass membership is renewed automatically every 12 months from the approval date. During the renewal process the Bank validates eligibility and informs the member of their complimentary visit entitlement by SMS before the expiry of the current membership. A new physical card is mailed when the membership is renewed, and the new expiry date is reflected in the Priority Pass account. To prevent disruption during the renewal period, the digital membership card should be used until the physical card arrives."]],

        ["What happens to a Priority Pass membership if the linked Standard Chartered card is cancelled?",
         ["If the Standard Chartered Priority Banking Visa Infinite Credit Card, Visa Infinite Credit Card or Journey Credit Card is cancelled, the Priority Pass membership will not be renewed, and the Bank will cease the usage of the Priority Pass membership with immediate effect without any notification."]],

        ["Can multiple Priority Pass memberships be linked to one Standard Chartered card?",
         ["No. Accounts cannot be linked, and a separate account is needed for each membership. All accounts are reflected in the Priority Pass mobile app with a different Priority Pass membership number, and each account must be logged in to with its respective membership number. Each Priority Pass membership should only be linked to the respective credit card used for registration."]],

        ["Can a different credit card be tagged to an existing Priority Pass membership?",
         ["No. Only the Standard Chartered Priority Banking Visa Infinite Credit Card, Visa Infinite Credit Card or Journey Credit Card used for registration can be linked to the Priority Pass membership. The membership will not be renewed if any other credit card is tagged to the Priority Pass account."]],

        ["Must Priority Pass credit card details be updated after a Standard Chartered card is replaced?",
         [["Yes. New credit card details must be updated in the Priority Pass account, as the card is used for any applicable charges. Failing to update the latest credit card details means the Priority Pass membership cannot be used to access the airport lounges. Details are updated by logging in at ", ["prioritypass.com/en/login", L.ppLogin], " under the “My Details” section."]]],

        ["Why are Priority Pass credit card details required by Standard Chartered?",
         ["Credit card details are required because applicable fees are charged directly to the Standard Chartered Priority Banking Visa Infinite Credit Card, Visa Infinite Credit Card or Journey Credit Card."]],
      ]},

      { name: "Priority Pass — Lounge Visits and Usage", url: U.prPass, qas: [
        ["How long is a Standard Chartered Priority Pass membership and its complimentary visits valid?",
         ["The Priority Pass membership and complimentary visits are valid for 12 months from the approval date of the Priority Pass membership."]],

        ["How many complimentary Priority Pass visits do Standard Chartered Priority Private customers receive?",
         ["Priority Private customers receive 24 airport lounge visits, subject to maintaining S$1,500,000 in monthly average balance on their Assets Under Management with the Bank in the last 12 months and maintaining active Wealth holdings with the Bank. With effect from 1 July 2026, if a principal cardholder is accompanied by a guest on each of 12 lounge visits, that counts as 24 visits and the complimentary visits are exhausted."]],

        ["How many complimentary Priority Pass visits do Qualified Standard Chartered Priority Banking customers receive?",
         ["With effect from 1 July 2026, Qualified Priority Banking customers receive 6 airport lounge visits, subject to maintaining S$200,000 in monthly average balance on their Assets Under Management with the Bank in the last 12 months and maintaining active Wealth holdings with the Bank. If a principal cardholder is accompanied by a guest on each of 3 lounge visits, that counts as 6 visits and the complimentary visits are exhausted. On or before 30 June 2026 the entitlement was 12 airport lounge visits."]],

        ["How many complimentary Priority Pass visits do other Standard Chartered Priority Banking customers receive?",
         ["Other Priority Banking customers receive 2 airport lounge visits, subject to maintaining active Wealth holdings with the Bank. If a principal cardholder is accompanied by a guest on a lounge visit, that counts as 2 visits and the complimentary visits are exhausted."]],

        ["What counts as active Wealth holdings for Standard Chartered Priority Pass entitlements?",
         ["Active Wealth holdings means customers who hold either Banca or Investment products with the Bank. Investment products are Structured Products, Fixed Income Products, Unit Trusts and Equity Holdings. Foreign Currency Deposits and FX transactions are not considered Investment Products. Customers who do not hold active Wealth holdings will have their benefits deactivated."]],

        ["How does a guest count towards Standard Chartered Priority Pass complimentary visits?",
         ["Each accompanying guest counts as one visit. For example, if a Principal Visa Infinite Cardholder is accompanied by two guests, that counts as three complimentary visits. Any additional visits by the principal cardholder or their guests beyond the complimentary entitlement will be chargeable."]],

        ["Can unused Standard Chartered Priority Pass complimentary visits be carried forward at renewal?",
         ["No. No complimentary visits can be carried over to the membership’s new visit plan after renewal. For example, it is not possible to bring forward 3 remaining complimentary visits from an old visit plan into a new plan that has 24 visits. All complimentary visits should be fully used before the membership expires."]],

        ["What happens after Standard Chartered Priority Pass complimentary visits are exhausted?",
         ["Priority Pass (A.P.) Limited charges for every subsequent visit after the allocated number of complimentary visits is exhausted. Such charges are as per Priority Pass (A.P.) Limited’s prevailing rates, which can be accessed via the Priority Pass website, and apply to both the cardholder and their guests."]],

        ["Why were Priority Banking Visa Infinite Credit Card lounge entitlements deactivated, and how are they reinstated?",
         ["As of 15 August 2025, active wealth holdings must be maintained with the Bank to continue enjoying the lounge entitlements. To reinstate the entitlements, wealth holdings must be purchased and maintained to meet the requirement, and the Bank reactivates the benefit within 6 to 8 weeks, at the end of the calendar month."]],

        ["Can an airport lounge be accessed without a Priority Pass card?",
         ["Lounge access is only available on presentation of a valid Priority Pass membership card. Not all Priority Pass lounges accept the digital membership card, and for those lounges the physical membership card must be produced. Standard Chartered recommends always carrying the physical membership card when travelling."]],

        ["How is a digital Priority Pass membership card used?",
         ["The Priority Pass digital membership card provides card-free access to the majority of Priority Pass lounges, using the 2D barcode found within the smartphone application. Not all Priority Pass lounges can be accessed with the digital membership card, so the physical Priority Pass membership card should also be carried when travelling."]],

        ["How is Priority Pass visit history and remaining entitlement checked?",
         ["The Priority Pass mobile app can be downloaded and logged in to, where the complimentary visit entitlement and visit history are shown under the “Account” section."]],

        ["What should be done about an incorrect Priority Pass guest lounge charge?",
         ["In the event of a dispute, a digital or physical copy of a signed receipt must be submitted as proof. The details on the digital screen or paper voucher presented at check-in to the lounge should be verified as accurate before endorsing it, and any errors highlighted to the lounge staff prior to entering the lounge."]],
      ]},
    ],
  },

  // ======================================================= FRAUD AND SCAMS
  {
    name: "Fraud and Scams",
    subs: [
      { name: "Fraud and Scam Overview", url: U.fsIndex, qas: [
        ["Which fraud and scam typologies does Standard Chartered Singapore cover?",
         ["Standard Chartered Singapore publishes guidance on the following typologies:",
          B([["Compromised device and malware attack."],
             ["Impersonation scam and phishing scam."],
             ["Investment scam and loan scam."],
             ["Money mule."]])]],

        ["What should a Standard Chartered client never disclose to anyone?",
         ["Personal particulars, Singpass and online banking credentials, credit and debit card details, and OTPs must never be disclosed to anyone."]],

        ["What should a Standard Chartered client do if they suspect unauthorised account access?",
         ["If any unauthorised access or transactions are suspected on an account, the mobile app or web browser should be closed immediately and the matter reported to the Bank without delay. Where login credentials may have been compromised, digital services can be temporarily disabled instantly."]],
      ]},

      { name: "Disable Your Digital Services", url: U.fsDisable, qas: [
        ["What does temporarily disabling Standard Chartered digital services do?",
         ["The feature temporarily disables Online Banking, Mobile Banking and Online Trading access immediately, for use whenever login credentials may have been compromised. It does not include blocking a compromised card or account. It is requested through the Client Contact Centre’s automated menu, option 9, which is the fastest channel, and can also be arranged through the live chat function, at any branch, or through a Relationship Manager where one is held."]],

        ["What are the steps to temporarily disable Standard Chartered Online and Mobile Banking access?",
         ["The steps are:",
          B([["Step 1: Reach the Client Contact Centre’s automated menu."],
             ["Step 2: Select option 9 in the main menu."],
             ["Step 3: Enter the 7-digit NRIC number or 16-digit debit or credit card number."],
             ["Step 4: Perform second level authentication via SC Mobile Push Notification or SMS OTP."],
             ["Step 5: Press 1 to confirm the request to disable Online and Mobile Banking and Online Trading login access."],
             ["An SMS notification is received upon successful deactivation of digital services."]])]],

        ["Can Online Trading still be used after Standard Chartered Online and Mobile Banking access is disabled?",
         ["No. Once Online and Mobile Banking is temporarily disabled, trading on the SC Online Trading platform is not possible."]],

        ["How is Standard Chartered Online and Mobile Banking access reactivated after being disabled?",
         ["Due to security reasons, reactivation of Online and Mobile Banking access can only be performed in person at any Standard Chartered branch. Supporting documents such as an NRIC or passport must be brought along for verification."]],

        ["Are cards and accounts blocked automatically when Standard Chartered digital services are disabled?",
         ["No. Debit and credit cards and accounts are not automatically blocked. Only Online Banking, Mobile Banking and SC Online Trading app access are temporarily blocked. Blocking a card or account is a separate request."]],

        ["Do scheduled transactions continue after Standard Chartered digital services are disabled?",
         ["Yes. All incoming and outgoing transactions scheduled prior to disabling Online and Mobile Banking access, such as scheduled transfers, GIRO and salary credit, continue to be processed."]],
      ]},

      { name: "Loan Scams", url: U.fsLoan, qas: [
        ["What is a loan scam?",
         ["Scammers claim to be licensed moneylenders and offer loan services to random users, often through SMS or WhatsApp messages. Interested parties are instructed to transfer money as a servicing fee before the loan can be disbursed. After making the transfer, victims find that the scammers are no longer contactable. As part of these scams, scammers may ask for personal information such as NRIC, Singpass details and bank account numbers, which is then used to harass or threaten victims for payment."]],

        ["How can the signs of a loan scam be spotted?",
         ["Scammers may reach out through the following methods:",
          B([["Advertisements claiming to provide instant, fuss-free loan approvals through social media, SMS, cold calls or messaging apps."],
             ["Fake business websites created to appear legitimate."],
             ["Requests for fees to be transferred before a loan application can be processed, and pressure to make quick decisions, for example special rates for today only or a promotion expiring."]])]],

        ["How can a person stay safe from loan scams?",
         ["Standard Chartered gives three tips:",
          B([["Check with the Bank if unsure, or on receiving calls or messages from someone claiming to be bank staff or an agent for the Bank. Seek financial help only from legitimate financial institutions registered with the ", ["Registry of Moneylenders", L.moneylenders], "."],
             ["Block and report the number on the platform where the message was received."],
             ["Never disclose personal particulars, banking and credit card details to anyone, especially over unsolicited phone calls."]])]],
      ]},
    ],
  },

  // ==================================================== REWARDS PROGRAMMES
  {
    name: "Rewards Programmes",
    subs: [
      { name: "Rewards Programmes Overview", url: U.rwIndex, qas: [
        ["Which rewards programmes does Standard Chartered Singapore offer?",
         ["Standard Chartered Singapore offers two rewards solutions:",
          B([["360° Rewards, which rewards credit card spends and eligible product holdings with the Bank."],
             ["360° Rewards Redemption, which redeems 360° Rewards points for e-vouchers, bill payments or conversion to miles."]])]],
      ]},

      { name: "360° Rewards Redemption Platform", url: U.rw360, qas: [
        ["What can be redeemed on the Standard Chartered 360° Rewards Redemption Platform?",
         ["The following redemptions are available:",
          B([["e-Vouchers of participating brands across retail, dining, lifestyle and travel. The e-Voucher details are retrieved directly on the platform for use."],
             ["Points to Cash, redeeming 360° Rewards points as cash credits."],
             ["Points to Miles, converting 360° Rewards points to miles."],
             ["Points Transfer, transferring 360° Rewards points to family members or friends. Points cannot be transferred from non-Visa Infinite cards to Visa Infinite cards."]])]],

        ["Why did Standard Chartered change its rewards redemption platform?",
         ["The platform was enhanced with an improved layout to provide a more convenient way to redeem 360° Rewards points, as part of working towards a more convenient banking experience."]],

        ["What is different on the new Standard Chartered rewards redemption platform?",
         ["The new platform offers improved processing lead time. Instead of physical mail, redemption codes are retrieved on the new rewards redemption platform directly, upon email confirmation sent to the email address updated with the Bank."]],

        ["What happens to redemptions made on the previous Standard Chartered rewards platform?",
         ["Any prior redemptions made are still processed. The statuses of these redemptions, and the past 3 months of redemption orders, can be viewed in the new rewards redemption platform."]],
      ]},
    ],
  },

  // =============================================== PROMOTIONS AND REFERRALS
  {
    name: "Promotions and Referrals",
    subs: [
      { name: "Promotions Overview", url: U.pmIndex, qas: [
        ["What promotions does Standard Chartered Singapore currently list?",
         ["The promotions page lists offers across the product range, including:",
          B([["Live The Good Life Privileges and SC Shop and Earn for credit cardholders."],
             ["e$aver, offering up to 1.60% p.a. interest from now till 30 September 2026 without locking in money."],
             ["Bonus$aver, offering interest of up to 5.85% p.a. and S$228 cashback on application."],
             ["Journey Credit Card, offering up to 45,000 miles."],
             ["Credit Card Funds Transfer, offering interest-free cash from as low as 1.80% processing fee (EIR from 4.86%)."],
             ["The Personal Loan Member-Get-Member Promotion, giving up to S$200 cashback for referring a friend to an eligible Standard Chartered Personal Loan."]])]],
      ]},

      { name: "Birthday Treats", url: U.pmBirthday, qas: [
        ["What birthday treats does Standard Chartered offer Priority clients?",
         ["Birthday treats available to Standard Chartered Priority clients include:",
          B([["LifeSpa: a spa treatment of choice and more, starting from S$20 for Priority Private clients and S$30 for Priority Banking clients, during the birthday month and the month after."],
             ["IHG Hotels and Resorts: 20% off stays at IHG properties across Asia-Pacific when paying with the Priority Banking Visa Infinite card."],
             ["Visa Luxury Hotel Collection: a premium collection of benefits at a selection of prestigious properties with the Priority Banking Visa Infinite card."],
             ["SC Shop and Earn: cashback at over 250 online merchants, with higher cashback for clients with a Priority relationship."]])]],
      ]},

      { name: "Caltex Fuel Savings", url: U.pmCaltex, qas: [
        ["What is the Standard Chartered and Caltex promotion?",
         ["From 1 January 2026 to 31 December 2026, cardholders can enjoy up to 27.91% fuel savings with a Simply Cash, Beyond or Visa Infinite Credit Card when paying on the CaltexGO App, or up to 24.19% fuel savings on an indoor payment at Caltex with a Simply Cash Credit Card. New and existing Standard Chartered cardholders can also get a S$10 instant discount for their first in-app fuel payment on CaltexGO, or S$30 in Caltex Fuel eVouchers when they sign up for a new credit card."]],

        ["How is a card registered for the Standard Chartered CaltexGO discount?",
         ["Registration in the CaltexGO app works as follows:",
          B([["Step 1: Tap the “My Account” icon from the home screen."],
             ["Step 2: Tap and add a payment method."],
             ["Step 3: Insert the preferred bank payment card."],
             ["Step 4: The discount amount shown depends on the card used."]])]],

        ["How do Standard Chartered CaltexGO fuel savings break down on a S$110 purchase?",
         ["On a S$110 fuel purchase paid through CaltexGO, the savings are:",
          B([["New to CaltexGO instant discount: S$10 off with a minimum spend of S$20."],
             ["Standard Chartered rebate: S$3 off with a minimum gross fuel spend of S$100."],
             ["Eligible upfront petrol discount, with no minimum spend required: 17%, which is S$16.49."],
             ["Simply Cash Credit Card only: an additional 1.5% cashback, which is S$1.21."],
             ["Total savings: S$30.70 or 27.91% with a Simply Cash Credit Card, and S$29.49 or 26.81% with a Beyond or Visa Infinite Credit Card."]])]],

        ["How do Standard Chartered Caltex indoor payment savings break down on a S$110 purchase?",
         ["On a S$110 Platinum 98 fuel purchase paid indoors at a Caltex station, the savings are:",
          B([["Caltex Fuel eVouchers: S$5 off, from 6 vouchers of S$5, with a minimum spend of S$70."],
             ["Standard Chartered rebate: S$3 off with a minimum gross fuel spend of S$100."],
             ["Eligible upfront petrol discount, with no minimum spend required: 17%, which is S$17.34."],
             ["Simply Cash Credit Card only: an additional 1.5% cashback, which is S$1.27."],
             ["Total savings: S$26.61 or 24.19% with a Simply Cash Credit Card, and S$25.34 or 23.04% with a Beyond or Visa Infinite Credit Card."]])]],
      ]},

      { name: "Instant Digital Credit Card", url: U.pmInstant, qas: [
        ["What is a Standard Chartered Digital Credit Card?",
         ["A Digital Credit Card is a virtual version of a credit card which can be used after successful activation without the physical card. The physical credit card is still received within 5 to 7 working days and must be activated before use. Once the physical credit card is activated, the Digital Credit Card is no longer available for use."]],

        ["What can a Standard Chartered Digital Credit Card be used for?",
         ["A Digital Credit Card can be used for:",
          B([["Online purchases using a one-time password with merchants that use 3-Domain Secure (3DS), an added layer of security for online credit card transactions. The OTP is sent to the principal cardholder’s mobile number registered with the Bank."],
             ["Digital wallets, namely Apple Pay, Google Pay and Samsung Pay, which allow the card to be enrolled and used to make purchases through a mobile device instead of the physical card."]])]],

        ["Who can get a Standard Chartered Digital Credit Card?",
         ["A Digital Credit Card may be made available to a principal cardholder in the following scenarios:",
          B([["On successfully applying for a Credit Card, including a Bonus$aver Credit Card, via the Bank’s roadshows, telesales or tele-marketing by the Bank, or any of the Bank’s branches."],
             ["On using MyInfo to successfully apply online or via the Mobile app for an Unlimited Cashback, Rewards+, Spree, NUS Alumni Platinum, Prudential Platinum, Visa Infinite or X Credit Card, where the card is instantly approved."],
             ["With effect from 2 September 2019 inclusive, when issued a replacement or renewal Credit Card."],
             ["A Digital Credit Card will not be issued for any Standard Chartered Platinum Mastercard Credit Card issued pursuant to a Debt Consolidation Plan application."]])]],

        ["By when must a Standard Chartered Digital Credit Card be activated?",
         ["A Digital Credit Card must be activated within 15 calendar days from the date the Credit Card application is approved, or from the date of issue of the replacement or renewal credit card. It cannot be activated after those 15 calendar days, or after the physical version of the Credit Card is activated, whichever is earlier."]],

        ["How is a Standard Chartered Digital Credit Card activated?",
         ["Activation depends on the card:",
          B([["Bonus$aver Credit Card, renewal credit card and replacement credit card: log in to Standard Chartered Online Banking or the Standard Chartered Mobile banking app to activate."],
             ["All other credit cards: either log in to Standard Chartered Online Banking or the Mobile banking app, or click the “Activate my Digital Credit Card” button on the Credit Card approval page for applications made online or via the SC Mobile App, where that page is made available by the Bank."]])]],

        ["What is MyInfo and how does it support a Standard Chartered application?",
         ["MyInfo is a service which allows SingPass users to manage their personal data and pre-fill forms in online transactions. It includes government-verified data retrieved across participating Government agencies and data the user contributed to form their profile. It can be used to apply for an instant Digital Credit Card or for disbursement of a CashOne Personal Loan or Credit Card Instalment Loan."]],
      ]},

      { name: "Million Reasons to be Happier", url: U.pmMillion, qas: [
        ["What is the Standard Chartered Million Reasons to be Happier promotion?",
         ["Million Reasons to be Happier is a promotion in which clients bank and spend with Standard Chartered for a chance to win a private jet trip to the Maldives. Applying for a Standard Chartered Credit Card is one way to participate. Terms and conditions apply."]],
      ]},

      { name: "SC Referral Club — Rewards and Eligibility", url: U.pmReferTnc, qas: [
        ["What is the Standard Chartered SC Referral Club Refer and Earn Promotion?",
         ["The SC Referral Club Refer and Earn Promotion is valid from 1 July 2026 to 30 September 2026, both dates inclusive. It is open to customers of Standard Chartered Bank (Singapore) Limited who hold a valid current, cheque or savings account as a primary account holder, or are principal cardholders of credit cards issued by the Bank, and who successfully refer a family member or friend to sign up for participating products during the Promotion Period."]],

        ["Which products participate in the Standard Chartered SC Referral Club promotion?",
         ["The Participating Products are the Visa Infinite Card and the Beyond Credit Card, each of which is an Eligible Card under the promotion."]],

        ["What rewards does a Standard Chartered SC Referral Club referrer receive?",
         ["Per successful referred individual, the referrer receives:",
          B([["Eligible Card other than the Beyond Credit Card: S$80."],
             ["Beyond Credit Card: 20,000 miles, credited in the form of 50,000 Rewards Points. Where the existing customer does not hold a rewards-points-earning Standard Chartered Card, they receive S$200 cashback instead."],
             ["There is no cap on the total amount of Referral Cashback or Referral Miles."]])]],

        ["When is a Standard Chartered SC Referral Club referral considered successful?",
         ["A referral is successful when all of the following are met:",
          B([["The existing customer shares their referral code, generated via the Standard Chartered Mobile App, during the Promotion Period."],
             ["The referred individual applies for at least 1 Eligible Card using that referral code during the Promotion Period."],
             ["The Bank receives the application within the Promotion Period. The Bank may continue to process applications for up to 14 calendar days after the end of the Promotion Period, at its sole and absolute discretion."],
             ["The application for the Eligible Card is approved by the Bank, and such approval is final and unconditional."],
             ["The referred individual applying for an Eligible Card other than the Beyond Credit Card is not an existing principal cardholder of any credit card issued by the Bank, and does not have any previously cancelled credit card in the last 12 months prior to the Referral Date."],
             ["The referred individual is at least 21 years old as at the Referral Date, which is the date they apply using the referral code."]])]],

        ["What must a Standard Chartered SC Referral Club referrer do before referring someone?",
         ["Before referring a referred individual, the existing customer must:",
          B([["Consent to the Bank disclosing their name, and the fact that they hold at least 1 deposit account or credit card issued by the Bank, to the referred individual."],
             ["Inform the referred individual that Referral Cashback or Referral Miles will be accorded to the referrer if the promotion terms are complied with."],
             ["Undertake to comply with the Personal Data Protection Act 2012 and all relevant subsidiary legislation, including obtaining any required consent from the referred individual to receive the referral code."],
             ["Inform and obtain the consent of the referred individual that the application status of their eligible account or card would be disclosed to the referrer by virtue of whether the referrer receives Referral Cashback or Referral Miles."]])]],

        ["When is Standard Chartered SC Referral Club cashback or miles credited?",
         ["Referral Cashback or Referral Miles are credited into one of the existing customer’s Standard Chartered deposit accounts or credit cards within 60 days from the date of approval of the referred individual’s Eligible Card, subject to that account or card being valid and in good standing at the time of crediting. The reward is forfeited if the account or card is suspended or closed, or if the existing customer is in breach of the banking agreement at the time of crediting. The Bank has the discretion to decide which account the reward is credited into."]],

        ["What happens if more than one person refers the same individual under the Standard Chartered SC Referral Club?",
         ["If more than one existing customer refers the same referred individual during the Promotion Period, only one existing customer is eligible for the Referral Cashback or Referral Miles. That is the existing customer whose referral code was used by the referred individual to make the application. In the event of any dispute, the Bank has the sole and absolute discretion to determine which existing customer is eligible."]],

        ["When can Standard Chartered decline to credit SC Referral Club rewards?",
         ["The Bank reserves the right to decline or cancel the credit of Referral Cashback or Referral Miles where it determines that any requirement or term of the banking agreement has not been complied with, including where:",
          B([["The existing customer has acted fraudulently or dishonestly."],
             ["The existing customer conducted themselves in bad faith or otherwise in an inappropriate manner to gain an unfair advantage against the Bank."],
             ["Any event occurs giving rise to a right for the Bank to suspend or terminate any of the banking agreement for a product."],
             ["The referred individual fails to meet the referral requirements, in which case the Bank may also substitute the reward with one of lower value."]])]],

        ["Are Standard Chartered SC Referral Club rewards transferable?",
         ["No. The Referral Cashback or Referral Miles are not exchangeable for any other items, and are non-transferable and non-assignable."]],
      ]},

      { name: "SC Referral Club — Signing Up with a Code", url: U.pmRefSignup, qas: [
        ["How is a Standard Chartered referral code used when signing up?",
         ["A referral code must be entered before applying for a product in order to receive referral rewards. Once the code is submitted successfully, the referred individual can begin applying for a product. The SC Referral Club promotion period runs from 1 July to 30 September 2026."]],

        ["Which cards can be applied for with a Standard Chartered referral code?",
         ["Referral promotions cover Priority Banking and Personal Banking products, including these credit cards:",
          B([["Visa Infinite Credit Card: earn up to 3 miles per S$1, with 6 complimentary Priority Pass lounge visits each year and up to 50,000 Welcome Miles when S$2,000 is spent within 60 days of card approval. An annual fee of S$550 excluding GST applies."],
             ["Beyond Credit Card: earn up to 8 miles per S$1, with complimentary Business Class upgrades for a limited time when one Business Class ticket and one Premium Economy ticket are purchased, up to two times per card year. An annual fee of S$1,500 excluding GST applies."]])]],
      ]},

      { name: "Priority Banking Referral Catalogue", url: U.pmCatalogue, qas: [
        ["What does the Standard Chartered Priority Banking Referral Programme catalogue cover?",
         ["The catalogue is shown to someone who has been referred to sign up for a Priority Banking relationship. It sets out the two steps to start the relationship — choosing and opening a current or savings account, then bringing in a minimum of S$200,000 fresh funds via FAST transfer — and lists the eligible accounts, which are Bonus$aver, e$aver and USD$aver."]],
      ]},

      { name: "The Good Life Privileges", url: U.pmGoodLife, qas: [
        ["What is the Standard Chartered The Good Life Programme?",
         ["The Good Life Programme is open to all cardholders of credit and debit cards issued by Standard Chartered Bank (Singapore) Limited. Eligible cardholders must charge transactions to their cards at the merchants participating in the Programme to redeem or receive privileges. Privileges include but are not limited to discounts, offers, promotions, prizes, gifts, complimentary items, vouchers, rebates and redemptions."]],

        ["Can The Good Life privileges be combined with other offers or transferred?",
         ["No. Privileges are strictly non-transferable and non-exchangeable, whether for cash, credit or otherwise, and cannot be used in conjunction with any other merchants’ discounts, privileges, promotions, discount schemes, loyalty programmes, discount cards or vouchers, unless otherwise stated."]],
      ]},
    ],
  },

  // ================================================ PARTNER CARD PROMOTIONS
  {
    name: "Partner Card Promotions",
    subs: [
      { name: "Grab Cards and CashOne Promotion", url: U.grab, qas: [
        ["Is the Standard Chartered Grab cards and CashOne promotion still running?",
         ["No. The promotion page states that the promotion has ended. It offered cashback for applying for a Standard Chartered Credit Card and CashOne Personal Loan, and the referral offers it carried closed on 30 November 2021 for credit cards and 31 December 2021 for personal loans."]],

        ["Which products were featured in the Standard Chartered Grab promotion?",
         ["The promotion featured three products:",
          B([["Unlimited Cashback Credit Card: 1.5% cashback on all spends, with no minimum spend and no cashback cap."],
             ["Rewards+ Credit Card: up to 10x rewards points on foreign currency for overseas retail, dining and travel, and up to 5x rewards points on dining transactions in SGD."],
             ["CashOne Personal Loan: an interest rate as low as 3.48% p.a. (EIR from 6.95% p.a.), with cash received in 15 minutes."]])]],
      ]},

      { name: "Qoo10 Cards and CashOne Promotion", url: U.qoo10, qas: [
        ["Is the Standard Chartered Qoo10 cards and CashOne promotion still running?",
         ["No. The promotion page states that the promotion has ended. It offered cashback for applying for a Standard Chartered Credit Card and CashOne Personal Loan, and carried the same referral offers that closed on 30 November 2021 for credit cards and 31 December 2021 for personal loans."]],
      ]},
    ],
  },

  // ============================================ BANK INFORMATION AND SERVICES
  {
    name: "Bank Information and Services",
    subs: [
      { name: "About Standard Chartered Singapore", url: U.about, qas: [
        ["What is Standard Chartered Bank’s presence in Singapore?",
         ["Standard Chartered Bank in Singapore is part of an international banking group with a presence in 55 of the world’s most dynamic markets. Its purpose is to drive commerce and prosperity through its unique diversity, and its heritage and values are expressed in the brand promise, here for good. The Bank has a history of 167 years in Singapore, where it opened its first branch in 1859. In October 1999 it was among the first international banks to receive a Qualifying Full Bank licence."]],

        ["What is Standard Chartered Bank (Singapore) Limited?",
         ["In 2013 the Bank transferred its Singapore Retail and SME businesses to a locally incorporated subsidiary, Standard Chartered Bank (Singapore) Limited (SCBSL). In May 2019 it fully consolidated its Singapore business operations through the transfer of its Commercial Banking, Corporate & Institutional Banking and Private Banking businesses to SCBSL. Singapore is home to the majority of the group’s global business leadership, its technology operations, and SC Ventures, its innovation hub."]],

        ["How is Standard Chartered Bank (Singapore) Limited rated?",
         ["SCBSL is well-rated with strong credit fundamentals: A1/Stable by Moody’s Investor Services, A+/Stable by Standard & Poor’s, and A+/Stable by Fitch Ratings. In August 2020 it was the first and only bank to be awarded Significantly Rooted Foreign Bank (SRFB) status by the Monetary Authority of Singapore, and in December 2020 it was granted enhanced SRFB privileges in recognition of a significantly higher degree of rootedness exceeding the SRFB baseline criteria."]],

        ["What is Trust Bank?",
         ["In September 2022 Standard Chartered launched Trust Bank, its digital bank venture in Singapore, in partnership with FairPrice Group. Trust Bank has become one of the world’s fastest-growing digital banks, expanding to over 1 million customers, equivalent to around 20% of the Singapore market, in just over 2 years since launch. It offers a wide range of products and services for retail customers, enabling them to save, spend, budget, borrow, invest and insure, and has received industry awards including best digital bank in Singapore by The Asian Banker and best mobile banking app globally by The Digital Banker."]],

        ["What is Standard Chartered Singapore’s branch and ATM network?",
         ["The Bank has a network of 11 branches, including 5 Priority Banking Centres and 1 International Banking and Priority Private Centre, and over 30 Standard Chartered ATMs."]],
      ]},

      { name: "International Banking", url: U.intl, qas: [
        ["What does Standard Chartered International Banking offer?",
         ["International Banking supports clients managing wealth across borders, with CIO-led insights and deep market expertise across Asia, Africa and the Middle East. Accounts can be opened remotely from anywhere in the world to start a new Priority Banking relationship, with up to USD 6,400 in rewards available. Terms and conditions apply."]],

        ["Why does Standard Chartered position Singapore as an international banking hub?",
         ["Standard Chartered cites three reasons:",
          B([["An attractive tax regime: individual income tax is progressive, and there is no capital gains, inheritance or estate duty. Corporate tax is flat, with incentives such as exemptions, rebates and deductions."],
             ["A top global business hub: Singapore ranks highly for ease of doing business, IP protection and low corruption, with Economic Development Board incentives and an efficient regulatory framework that simplifies incorporation."],
             ["Access to the world: Singapore’s strategic location in Southeast Asia makes its port a key hub, and its 25 free trade agreements with major markets give investors and businesses access to important global economies."]])]],

        ["What is Standard Chartered’s Global Chinese service?",
         ["Global Chinese services support the journeys of Chinese clients around the world who are looking to bank internationally, providing support across the Bank’s global network."]],
      ]},

      { name: "Business Banking", url: U.business, qas: [
        ["What does Standard Chartered Singapore SME Banking offer?",
         ["Standard Chartered SME Banking is positioned to make banking easy, quick and transparent for small and medium enterprises in Singapore, providing business banking products and services alongside partner offerings."]],
      ]},

      { name: "Personal Banking", url: U.personal, qas: [
        ["What does the Standard Chartered Singapore personal banking page provide?",
         ["The personal banking page is a product finder that helps a client identify the product that best suits their banking needs, spanning the Bank’s deposit, credit card, lending, investment and insurance ranges."]],
      ]},

      { name: "SC Inner Circle", url: U.inner, qas: [
        ["What is SC Inner Circle?",
         ["SC Inner Circle is a Standard Chartered community in which clients and the Bank come together not as bankers and clients but as co-creators, collaborating to provide a best-in-class banking experience. Members’ insights are used to challenge the Bank’s leading-edge technologies, AI and security. The community is 12,000 members strong."]],

        ["What do SC Inner Circle members take part in?",
         ["Members receive invites to exclusive events and webinars, and participate in closed-door discussions, video calls and surveys to share ideas and inputs. Past activities have included a learning journey with Nanyang Polytechnic students, a Gratitude Jar initiative, Customer Service Week, and volunteering at Bright Hill Evergreen Home."]],

        ["How does a client join SC Inner Circle?",
         ["A client registers their interest by providing their full name and email address and agreeing to the Terms of Participation. The Terms of Participation should be read before registering, and members also agree that the Bank may contact them for SC Inner Circle events or surveys."]],
      ]},

      { name: "SC eSign", url: U.esign, qas: [
        ["What is Standard Chartered eSign?",
         ["Standard Chartered eSign allows documents to be signed electronically via a mobile device or computer."]],

        ["Which documents can be signed with Standard Chartered eSign?",
         ["The following document types can be signed with eSign:",
          B([["Mortgage documents and agreements, and Mortgage Repricing."],
             ["Insurance."],
             ["Referral Agreement."],
             ["Priority Banking Servicing Forms."],
             ["Tax Residency Declaration."]])]],

        ["What is needed to use Standard Chartered eSign?",
         ["The user selects what they are signing for and the relevant account, then enters their email address and their unique eSign code before submitting."]],
      ]},

      { name: "SC Digital KYC Review", url: U.cdd, qas: [
        ["What is the Standard Chartered Digital KYC Review portal?",
         ["The Standard Chartered KYC Review portal allows a KYC review to be completed at the user’s convenience, helping fight financial crime and fraud by keeping information updated."]],

        ["How is a Standard Chartered Digital KYC Review accessed?",
         ["The user enters the unique code sent to their registered email address, and confirms that they are authorised by the entity or party they represent to provide the details, information and documents necessary for KYC review on behalf of that entity or party. On verification of the unique code, a one-time password is sent to the registered mobile number. Standard Chartered accepts any user of the portal as being authorised to act on behalf of the relevant entity or party once the unique code and one-time password are verified, and is not required to check that person’s authority."]],
      ]},

      { name: "Statements Refresh", url: U.statements, qas: [
        ["What has changed in Standard Chartered Singapore statements?",
         ["Standard Chartered has refreshed its statements and published a guide to the new format. Guides are available for Consolidated Statements, Loans, Credit Cards and Personal Loans, and Savings and Current Accounts."]],

        ["What sections does a Standard Chartered Consolidated Statement contain?",
         ["A Consolidated Statement covers account balances plus account activities for the following:",
          B([["Commodity Account (gold and silver), and Securities Trading Account."],
             ["Investment Funds and Portfolio Account, and Bonds."],
             ["Structured Notes, including Structured Notes – Principal at Risk."],
             ["Commodity Linked Structured Investment (CLSI) and Premium Currency Investment (PCI)."],
             ["Structured Deposits (SD)."],
             ["Life Insurance and Investment Linked Products (LI / ILP)."]])]],
      ]},

      { name: "Journey Credit Card Rebrand", url: U.journeyFaq, qas: [
        ["What happened to the Standard Chartered X Credit Card on 19 May 2023?",
         ["With effect from 19 May 2023, the X Credit Card was rebranded to the Journey Credit Card with a refreshed card design. Existing X Credit Card holders could continue to use their card until it was due for renewal. The card programme switched from Visa Infinite to Visa Signature, and cardholders can earn up to 3 miles for each dollar spent on bonus transactions, up to a cap of 3,000 miles per statement month. These enhancements were applied automatically from 19 May 2023."]],

        ["What happened to fees when the X Credit Card became the Journey Credit Card?",
         ["The annual fee on the X Credit Card was S$702 including 8% GST. The annual fee on the Journey Credit Card was reduced to S$194.40 including 8% GST. All other fees and charges remained the same."]],

        ["What happened to instalment plans and balances on the X Credit Card at rebrand?",
         ["There was no change to existing instalment plans or payment arrangements, including recurring payment plans, instalment plans, EasyPay on Transactions and EasyPay on Balances. All payment arrangements were transferred to the new Journey Credit Card automatically, and balances on the X Credit Card were also transferred automatically."]],

        ["What happened to 360° Rewards Points and supplementary cards at the X Credit Card rebrand?",
         ["There was no change to existing 360° Rewards Points earned on the X Credit Card, and points earned on the Journey Credit Card from 19 May 2023 have no expiry date. Supplementary X Credit Cards were replaced together with the principal credit card."]],

        ["How does the Journey Credit Card bonus cap work on a S$1,200 bonus category spend?",
         ["Bonus category spend earns a maximum of 3,000 miles, credited in equivalent 360° Rewards Points, per statement cycle. Standard Chartered illustrates this as follows:",
          B([["S$1,000 spent at NTUC FairPrice Online: 3,000 base points plus 4,500 additional points, totalling 7,500 points, equivalent to 3,000 KrisFlyer miles."],
             ["S$1,200 spent at NTUC FairPrice Online: 3,600 base points plus 4,500 additional points capped, totalling 8,100 points, equivalent to 3,240 KrisFlyer miles. Only 4,500 additional points are awarded because the Journey Cap has been met."]])]],

        ["What is a carbon neutral certified card?",
         ["The Journey Credit Card is certified carbon neutral, which is achieved by calculating the carbon footprint in the card’s manufacturing process and reducing that to zero. The card is printed on carbon neutral plastic, with the 16-digit card number and expiry date printed at the back of the card."]],
      ]},

      { name: "Allianz Hospital Income Protect", url: U.allianz, qas: [
        ["What is Allianz Hospital Income Protect?",
         ["Allianz Hospital Income Protect is a hospital insurance plan distributed by Standard Chartered and underwritten by Allianz, providing cover for hospitalisation expenses. A promotion offering 10% off the total first year premium applies to plans purchased by 30 September 2026, subject to terms and conditions."]],

        ["What benefits does Allianz Hospital Income Protect provide?",
         ["The plan provides the following benefits, based on the Allianz Accident Protect Platinum Plan and in SGD:",
          B([["Infectious Disease: a lump sum payout of up to $600 upon diagnosis of infectious diseases including Hand Foot Mouth Disease and dengue."],
             ["Daily Hospitalisation Cash Benefit: daily hospitalisation income, as a result of injury or illness including Covid-19, for up to 750 days. The payout doubles for a stay in an intensive care unit for up to 90 days."],
             ["Recuperation Benefit: an additional daily cash benefit of up to $300 paid during the medical leave period."],
             ["Terms, conditions and exclusions apply."]])]],
      ]},
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Others FAQ",
  docTitle: "SCB Bank_Others FAQ",
  description: "Standard Chartered Singapore Others FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
console.log("categories:", CATEGORIES.length, "subcategories:", CATEGORIES.reduce((n, c) => n + c.subs.length, 0));
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Others_FAQ.docx", buf);
  console.log("written", buf.length);
});
