const fs = require("fs");
const { Packer } = require("docx");
const { buildDocument, B } = require("./faq_doc_builder");

const BASE = "https://www.sc.com/sg/bank-with-us/";
const U = {
  updateSignature: BASE + "service-requests/update-signature/",
  statementRequest: BASE + "service-requests/statement-request/",
  suppCardUpdate: BASE + "service-requests/other-requests/suppcardupdate/",
  creditLimitDecrease: BASE + "service-requests/other-requests/credit-limit-decrease/",
  ccil: BASE + "service-requests/other-requests/ccil/",
  mortgageRepricing: BASE + "service-requests/mortgage-repricing/",
  linkDelink: BASE + "service-requests/link-delink-accountondebitandcreditcard/",
  embossName: BASE + "service-requests/emboss-name-change/",
  earlyCcip: BASE + "service-requests/early-ccip-payment/",
  dailyLimit: BASE + "service-requests/changedailylimitfordebitorcreditcards/",
  statementCycle: BASE + "service-requests/cc-statement-cycle-date-change/",
  dispUnauth: BASE + "service-requests/card-transaction-dispute/unauthorized-transaction/",
  dispGoods: BASE + "service-requests/card-transaction-dispute/goods-and-or-services-issues/",
  dispError: BASE + "service-requests/card-transaction-dispute/error-in-transaction-details/",
  dispAtm: BASE + "service-requests/card-transaction-dispute/atm-cash-withdrawal-issues/",
  cardReplacement: BASE + "service-requests/card-replacement/",
  secToken: BASE + "security/security-token-card/",
  secPractices: BASE + "security/online-banking-security-practices/",
  secGuarantee: BASE + "security/online-banking-security-guarantee/",
  secFeatures: BASE + "security/online-banking-security-features/",
  myfIndex: BASE + "manage-your-finances/",
  clr: BASE + "manage-your-finances/credit-limit-review/",
  easybill: BASE + "manage-your-finances/easybill/",
  easypay: BASE + "manage-your-finances/easypay/",
  ipp: BASE + "manage-your-finances/ipp/",
  visaIpp: BASE + "manage-your-finances/ipp/visainstalments/",
  tcli: BASE + "manage-your-finances/temporary-credit-limit-increase/",
  mypIndex: BASE + "manage-your-payments/",
  applePay: BASE + "manage-your-payments/apple-pay/",
  googlePay: BASE + "manage-your-payments/google-pay/",
  samsungPay: BASE + "manage-your-payments/samsung-pay/",
  scMobile: BASE + "mobile-banking-services/standard-chartered-mobile/",
  huawei: BASE + "mobile-banking-services/standard-chartered-mobile/huawei-appgallery/",
  billMerchants: BASE + "online-banking-services/online-bill-payment-merchants/",
  eStatements: BASE + "online-banking-services/online-estatements/",
  onlineBanking: BASE + "online-banking-2/",
  alerts: BASE + "online-banking/alerts-and-notifications/",
  fees: BASE + "understanding-fees-and-charges/",
  latePayment: BASE + "understanding-late-payment-charge/",
  remittance: BASE + "remittance/",
  moneyLock: BASE + "money-lock/",
  stayHome: BASE + "stay-home/",
  hereForYou: BASE + "here-for-you/",
  digitalBanking: BASE + "digital-banking/",
};

const L = {
  pricingGuide: "https://www.sc.com/sg/pricing-guide/",
  clrTnc: "https://www.sc.com/sg/terms-and-conditions/credit-limit-review/",
  tcliTnc: "https://www.sc.com/sg/terms-and-conditions/temporary-credit-limit-increase/",
  easypayTxnTnc: "https://www.sc.com/sg/terms-and-conditions/easypay-on-transaction/",
  easypayBalTnc: "https://www.sc.com/sg/terms-and-conditions/easypay-on-retail-balances/",
  visaIpp: BASE + "manage-your-finances/ipp/sc-visa-instalment-payment-plan/",
  billMerchants: BASE + "online-banking-services/online-bill-payment-merchants/",
};

const CATEGORIES = [
  // ======================================================= SERVICE REQUESTS
  {
    name: "Service Requests",
    subs: [
      { name: "Update Signature", url: U.updateSignature, qas: [
        ["How is a signature updated on a Standard Chartered Singapore account?",
         ["A signature is updated through SC Mobile as follows:",
          B([["Step 1: From Help & Services, select “Update My Signature”."],
             ["Step 2: Select the accounts to link the updated signature to."],
             ["Step 3: Follow the on-screen instructions to upload or capture the new signature."],
             ["Step 4: Review the change instructions and confirm by entering the OTP on the following screen."],
             ["Step 5: A submission confirmation appears on successful completion of the service request."]])]],
      ]},
      { name: "Statement Request", url: U.statementRequest, qas: [
        ["How is a past statement requested from Standard Chartered Singapore?",
         ["A statement request is raised through Help & Services as follows:",
          B([["Step 1: In Help & Services, “Statement Request” is found under Account Management, Card Management, Loan Management and Other banking requests."],
             ["Step 2: On the account, card or loan selection screen, only one account can be selected."],
             ["Step 3: Fill out the mandatory sections and choose between eStatement or Paper Statement."],
             ["Step 4: Select the account for debiting the charges for the statement request."],
             ["Step 5: A confirmation screen displays the request details."],
             ["Step 6: An acknowledgement screen confirms the submission."]])]],
      ]},
      { name: "Supplementary Card Update", url: U.suppCardUpdate, qas: [
        ["How are supplementary card details updated using the Other Requests feature?",
         ["Supplementary card details are updated through the Other Requests feature as follows:",
          B([["Step 1: Go to Help and Services > Other Banking Services > Other Request."],
             ["Step 2: Input the type of change required (for example, mobile number), the 16-digit credit card number, and the supplementary cardholder’s name as per NRIC or passport, then tap Next."],
             ["Step 3: Review the request and tap Next."],
             ["Step 4: The request is submitted and a reference number is displayed."]])]],

        ["How is the outcome of a supplementary card update request checked?",
         [["To find out the outcome or progress of a supplementary card update request submitted through Other Requests, log in to Online Banking and go to “Mailbox”."]]],
      ]},
      { name: "Credit Limit Decrease", url: U.creditLimitDecrease, qas: [
        ["How is a credit card limit lowered at Standard Chartered Singapore?",
         ["A credit limit decrease is requested through the Other Requests feature as follows:",
          B([["Step 1: Go to Help and Services > Other Banking Services > Other Request."],
             ["Step 2: Input “Credit limit decrease”, the 16-digit credit card number and the preferred credit limit in the text box, then tap Next."],
             ["Step 3: Review the request and tap Next."],
             ["Step 4: The request is submitted and a reference number is displayed. The outcome or progress can be checked by logging in to Online Banking and going to “Mailbox”."]])]],
      ]},
      { name: "CashOne / Credit Card Instalment Loan Tenure Change", url: U.ccil, qas: [
        ["How is the tenure of a CashOne or Credit Card Instalment Loan changed?",
         ["The loan tenure is changed through the Other Requests feature as follows:",
          B([["Step 1: Go to Help and Services > Other Banking Services > Other Request."],
             ["Step 2: Input “Change of Loan Tenure”, the CashOne or CCIL account number in the format 9702-2XXX-XXXX-XXXX, and the requested remaining loan tenure in months (for example 12, 24, 36 or 48), then tap Next."],
             ["Step 3: Review the request and tap Next."],
             ["Step 4: The request is submitted and a reference number is displayed. The outcome or progress can be checked by logging in to Online Banking and going to “Mailbox”."]])]],

        ["What conditions apply to changing a CashOne or Credit Card Instalment Loan tenure?",
         ["The following conditions apply to a change of loan tenure:",
          B([["There is a S$50 administrative charge for this change, which is deducted from the credit card balance."],
             ["The minimum remaining loan tenure must be at least 12 months."],
             ["At least one loan instalment must have been billed before a change in loan can be applied for."],
             ["The number of instalments paid so far added to the number of new instalments cannot exceed 60 months."],
             ["There must be no outstanding unpaid loan instalments."]])]],
      ]},
      { name: "Mortgage Repricing", url: U.mortgageRepricing, qas: [
        ["How is a mortgage repricing request submitted to Standard Chartered Singapore?",
         ["Mortgage repricing is requested through the Mortgage Management feature as follows:",
          B([["In Help & Services, scroll to “Service request by category” and under “Mortgage Management” select “Mortgage Repricing”."],
             ["Select the current Mortgage Loan account for repricing."],
             ["Review the request."],
             ["Submit the request."],
             ["View the status of the request."]])]],
      ]},
      { name: "Link or Delink Account on Debit and Credit Card", url: U.linkDelink, qas: [
        ["How is an account linked to or delinked from a Standard Chartered debit or credit card?",
         ["An account is linked or delinked as follows:",
          B([["Step 1: In Help & Services, scroll to “Service request by category” and under “Card Management” select “Link/delink Your Account to Debit Card” or “Link/delink Your Account to Credit Card”."],
             ["Step 2: Select the Debit Card or Credit Card to link or delink. Only 1 Debit Card or Credit Card can be selected per submission."],
             ["Step 3: Enter the request details. Important information to include is the preferred embossed name on the Debit/ATM card."],
             ["Step 4: Review the request details."],
             ["Step 5: The submission generates a reference number."]])]],
      ]},
      { name: "Emboss Name Change", url: U.embossName, qas: [
        ["How is the embossed name on a Standard Chartered credit or debit card changed?",
         ["The embossed name is changed as follows:",
          B([["Step 1: From Help & Services, select “Credit Card Emboss Name Change” or “Debit/ATM Card Emboss Name Change” accordingly."],
             ["Step 2: Select the credit or debit card whose embossed name is to be changed."],
             ["Step 3: Specify the name to change to and any other instructions."],
             ["Step 4: Review the request details."],
             ["Step 5: A submission confirmation appears on successful completion of the service request."]])]],
      ]},
      { name: "Early Credit Card Instalment Plan Payment", url: U.earlyCcip, qas: [
        ["How is early payment made for a Credit Card Instalment Plan (EasyPay)?",
         ["Early payment for a Credit Card Instalment Plan is requested as follows:",
          B([["Step 1: From Help & Services, select “Early Payment of Credit Card Instalment Plan”."],
             ["Step 2: Select the current credit card for early payment. Only 1 credit card can be selected per submission."],
             ["Step 3: Enter the request details, including a description of the instalment plan and the exact monthly instalment amount billed."],
             ["Step 4: Review the request details."],
             ["Step 5: A submission confirmation appears on successful completion of the service request."]])]],
      ]},
      { name: "Change Daily Limit for Debit or Credit Cards", url: U.dailyLimit, qas: [
        ["How is the daily limit changed on a Standard Chartered debit, ATM or credit card?",
         ["The daily limit is changed as follows:",
          B([["Step 1: In Help & Services, navigate to “Service request by category” and under “Card Management” select “Change Daily Limit for Debit / ATM / Credit Card”."],
             ["Step 2: Select the card type."],
             ["Step 3: Select the eligible card displayed."],
             ["Step 4: A pop-up shows the POS limit available."],
             ["Step 5: Select the POS limit and click Done."],
             ["Step 6: Check the request details on the confirmation screen."],
             ["Step 7: Complete OTP authentication."],
             ["Step 8: Submit the request."]])]],

        ["What is the default daily ATM limit on a Standard Chartered card?",
         [["The daily ATM limit is set to a default of SGD4,000. Non-SGD cards are set to equivalent SGD values."]]],
      ]},
      { name: "Credit Card Statement Cycle Date Change", url: U.statementCycle, qas: [
        ["How is the statement cycle date changed on a Standard Chartered credit card?",
         ["The credit card statement cycle date is changed as follows:",
          B([["Step 1: In Help & Services, scroll to “Service request by category” and under “Card Management” select “Credit Card Statement Cycle Date Change”."],
             ["Step 2: Select the current credit card for the statement cycle date change. Only 1 credit card can be selected per submission."],
             ["Step 3: Enter the request details, limited to 500 characters."],
             ["Step 4: Review the request details."],
             ["Step 5: The submission generates a reference number."]])]],

        ["Which dates can be selected as a new credit card statement cycle date?",
         [["A new credit card statement cycle date only starts from the 3rd to the 27th of the month."]]],
      ]},
      { name: "Dispute — Unauthorised Transaction", url: U.dispUnauth, qas: [
        ["How is an unauthorised card transaction disputed at Standard Chartered Singapore?",
         ["An unauthorised transaction dispute is raised as follows:",
          B([["Step 1: There are 2 issue options under “I didn’t authorize this transaction”. Each option has the same process flow."],
             ["Step 2: On selecting an issue option, a prompt to block the card is displayed. When selected, the card is permanently blocked and replaced once the entire request is submitted."],
             ["Step 3: Select a maximum of 10 transactions with issues."],
             ["Step 4: A police report is required for processing."],
             ["Step 5: Upload the police report and other evidence, and optionally write a description of the issue."],
             ["Step 6: To upload the police report and evidence, either take a photo or upload from photo or file folders, depending on device options."],
             ["Step 7: Enter a description of the issue, limited to 500 characters, and tick the declaration that the information is correct and accurate."],
             ["Step 8: Review the confirmation screen of all details before submission. The card will be blocked after submission and a police report must be filed."],
             ["Step 9: Submit the request and receive confirmation with a service number."],
             ["Step 10: The request status can be checked via the Status tab."]])]],

        ["Is a police report required to dispute an unauthorised card transaction at Standard Chartered Singapore?",
         [["Yes. A police report is required for processing an unauthorised transaction dispute. If the police report is not available at the time of submission, it may be provided within 1 week from the submission date of the dispute."]]],
      ]},
      { name: "Dispute — Goods and Services Issues", url: U.dispGoods, qas: [
        ["How is a dispute raised over goods or services on a Standard Chartered card transaction?",
         ["A goods or services dispute is raised as follows:",
          B([["Step 1: There are 3 issue options under “I have a problem with the goods and/or services”. Each option has the same process flow."],
             ["Step 2: Select a maximum of 10 transactions with issues."],
             ["Step 3: Once the transactions are selected, the next page displays the selected transaction issue and requests documents and a description."],
             ["Step 4: At the bottom of the page, merchant correspondence can be uploaded. A description of what was purchased and the expected delivery or services must be completed. A description of the issue is optional."],
             ["Step 5: Confirm all details on the confirmation screen."],
             ["Step 6: Submit the request and receive a service number."]])]],
      ]},
      { name: "Dispute — Error in Transaction Details", url: U.dispError, qas: [
        ["How is an error in card transaction details disputed at Standard Chartered Singapore?",
         ["A transaction details dispute is raised as follows:",
          B([["Step 1: There are 6 issue options under “I have a problem with the transaction details”. Each option has the same process flow."],
             ["Step 2: After selecting an issue option, up to 10 transactions with issues can be selected."],
             ["Step 3: Select transactions using the radio buttons."],
             ["Step 4: Once transactions are selected and Next is clicked, different transaction issues are displayed based on the option selected earlier. Scroll down on this page."],
             ["Step 5: At the bottom of the page, sales documents and merchant correspondence can be uploaded. A description of what was purchased is a mandatory field."],
             ["Step 6: To upload sales documents and merchant correspondence, either take a photo or upload from photo or file folders, depending on device options."],
             ["Step 7: Enter a description of what was purchased and an issue description, limited to 500 characters, and tick the declaration of correct and accurate information."],
             ["Step 8: Review the confirmation screen of all details before submission."],
             ["Step 9: Submit the request and receive confirmation with a service number."],
             ["Step 10: The request status can be checked via the Status tab."]])]],

        ["How far back can transactions be selected when disputing transaction details at Standard Chartered Singapore?",
         [["Only up to 90 days of transactions are shown when selecting transactions to dispute under “I have a problem with the transaction details”."]]],
      ]},
      { name: "Dispute — ATM Cash Withdrawal Issues", url: U.dispAtm, qas: [
        ["How is a problem with an ATM cash withdrawal reported to Standard Chartered Singapore?",
         ["An ATM cash withdrawal dispute is raised as follows:",
          B([["Step 1: There is 1 issue option under “I have a problem with an ATM cash withdrawal”."],
             ["Step 2: On selecting the issue option, select a maximum of 10 transactions with issues."],
             ["Step 3: Once the transactions are selected, the next page displays the selected transaction issue. The time and date of withdrawal, country and location must be entered."],
             ["Step 4: A description of the issue is optional. A checkbox must be ticked to confirm the information submitted is correct."],
             ["Step 5: Review the confirmation screen of all details before submission."],
             ["Step 6: Submit the request and receive confirmation with a service number."]])]],
      ]},
      { name: "Card Replacement", url: U.cardReplacement, qas: [
        ["How is a replacement card requested from Standard Chartered Singapore?",
         ["A replacement card is requested as follows:",
          B([["Step 1: Go to “Help & Services” in the Services tab at the bottom right corner, and tap “View All”."],
             ["Step 2: Under Card Management, select “Replace Card”."],
             ["Step 3: Select the cards to be replaced."],
             ["Step 4: Review the details."],
             ["Step 5: Submit the request. An email and SMS are sent when the request is completed."]])]],

        ["Is there a charge for a Standard Chartered card replacement?",
         [["No, the card replacement is free."]]],

        ["Does a Standard Chartered replacement card carry the same card number?",
         [["A replacement card is sent with the same card number. However, the CVV number on the back of the card will be new. Card details should be updated for any online credit card arrangements held with merchants, as well as mobile wallets such as Google Pay, Apple Pay and Samsung Pay."]]],

        ["Does a Standard Chartered replacement card need to be activated?",
         [["Yes, the replacement card must be activated before usage."]]],

        ["Do linked accounts need to be relinked to a Standard Chartered replacement card?",
         [["No. Existing account linkages remain on the replacement card, and accounts do not need to be relinked to the new card."]]],

        ["How is the status of a Standard Chartered service request checked?",
         [["The status can be checked by going to the “Status” tab under Help & Services on both Online Banking and Standard Chartered Mobile."]]],

        ["How long is the status of a Standard Chartered service request displayed?",
         [["“Active” requests are displayed until the status is updated to “Completed”. “Completed” requests are displayed for 90 days."]]],

        ["Is there an alert when a Standard Chartered service request is completed?",
         [["Yes. An SMS and an email are sent when the request is completed."]]],
      ]},
    ],
  },

  // =============================================================== SECURITY
  {
    name: "Security",
    subs: [
      { name: "Security Token Card", url: U.secToken, qas: [
        ["What is the Standard Chartered Security Token Card?",
         ["The Security Token Card is as slim and lightweight to carry as any other card and has 2-in-1 functionality. The card acts as both:",
          B([["A credit, debit or ATM card, depending on the type of card chosen to host the security token."],
             ["A security token to authenticate certain Online Banking activities."]])]],

        ["Which card types are eligible for the Standard Chartered Security Token Card?",
         ["Eligible card types are:",
          B([["Credit cards: Visa Platinum, MasterCard Platinum, Visa Infinite, Bonus$aver."],
             ["Debit cards: Super Salary, XtraSaver, Bonus$aver."],
             ["ATM cards and Priority Banking ATM cards."],
             ["Clients who do not hold any of the eligible credit, debit or ATM cards are issued a generic Token card with only the security token functionality."]])]],

        ["Why was the Standard Chartered Token Card introduced?",
         ["The Token Card was introduced as part of an industry-wide initiative to enhance online security, using a more sophisticated authentication process called Transaction Signing. From January 2013, the Token Card is required for the following key Online Banking activities:",
          B([["Adding third-party payees."],
             ["Transferring funds or making payments for amounts above a defined threshold."],
             ["Updating personal details such as mailing address, contact number and email address."]])]],

        ["What is Transaction Signing on the Standard Chartered Token Card?",
         [["Transaction Signing is an additional authentication capability. Transaction-specific information, such as the last 4 digits of a payee’s account number, is first entered into the Token Card. A 6-digit transaction signing PIN is then generated for input to “sign” or validate the transaction on Online Banking or mobile banking."]]],

        ["How is a Standard Chartered Token Card applied for?",
         [["To apply for a Token Card, log in to Online Banking, where a prompt to select and apply for the Token Card is presented. Clients without an Online Banking account can register for one first."]]],

        ["What terms apply to the Standard Chartered Token Card?",
         ["The following terms apply to the Token Card:",
          B([["There is no change in the credit limit of the selected credit card, or in the applicable fees and charges of the selected credit, debit or ATM card."],
             ["The credit card expiry date is renewed upon issuance of the new Token Credit Card to the principal cardholder. Supplementary cardholders are therefore also issued new supplementary credit cards, without the security token function, due to the corresponding change in credit card expiry date."],
             ["The card number of the selected credit card remains the same and there is no impact on existing payment arrangements set up on the credit card."],
             ["The card functionality on the new credit, debit or ATM Token Card must be activated before use."],
             ["Customer Terms, Electronic Banking Terms and Conditions, and Credit Card Terms apply."]])]],
      ]},
      { name: "Online Banking Security Practices", url: U.secPractices, qas: [
        ["What should a Standard Chartered client do to protect their Online Banking PIN?",
         ["Standard Chartered advises the following to protect an Online Banking PIN:",
          B([["Keep the PIN confidential at all times, memorise it and destroy the PIN advice."],
             ["Change the PIN immediately if it is suspected to have been revealed, and change it regularly through the “Personal Update – Change PIN” function in Online Banking."],
             ["Always log off from an online session when leaving the computer unattended, even briefly."],
             ["Use different PINs for different financial or non-financial web-based services such as email, online shopping, digital identity and other online subscription services."]])]],

        ["What should a Standard Chartered client avoid doing with their Online Banking PIN?",
         ["Standard Chartered advises against the following:",
          B([["Do not divulge the PIN to anyone at any point in time, and do not write it down or store it anywhere."],
             ["Do not use the PIN when others can observe."],
             ["Do not use sequential numbers such as 12345, or the same digit more than twice such as 12322, and do not recycle PINs."],
             ["Do not store the username or PIN in the Internet Explorer browser."],
             ["Do not use public or Internet cafe computers for online banking."],
             ["Do not choose a PIN that is easily identifiable, such as the username, personal telephone number, birthday or other personal information."],
             ["Standard Chartered will never contact a client directly to ask them to disclose their PIN or password information."]])]],

        ["What is Second Factor Authentication on Standard Chartered Online Banking?",
         [["Second Factor Authentication (2FA) is an additional layer of security used by the Bank to verify a customer’s online identity. Standard Chartered Online Banking customers must provide a unique Additional Logon PIN (ALP) to access personal account details and perform online transactions. This is required in addition to the existing Online Banking username and PIN, in line with recommendations issued by the Monetary Authority of Singapore for all banks to adopt two-factor authentication. The ALP can be sent to the customer by SMS or from a Security Device (Token) issued by Standard Chartered."]]],

        ["How can a client check that a Standard Chartered Online Banking session is secure?",
         ["Standard Chartered advises the following checks and practices when banking online:",
          B([["Always check that Standard Chartered Online Banking originates from the Standard Chartered public web site."],
             ["Check for the use of https:// in the URL. Some browsers change the colour of the URL window during a secure session."],
             ["Check for a digital certificate represented by a padlock or key. Double-clicking this icon provides information about the organisation with which the secure session has been entered."],
             ["Always log off completely from an Online Banking session. Closing the window may not close the banking session, and a session may be hijacked if the computer is infected with a Trojan."],
             ["Clear the browser’s cache and history after each Online Banking session."]])]],

        ["What should a Standard Chartered client avoid when banking online?",
         ["Standard Chartered advises against the following:",
          B([["Never click a link in an email to reach a website and enter personal details, either in the email or on the website."],
             ["Do not disclose personal information to senders of convincing emails offering the chance to make easy money."],
             ["Do not disclose personal details or card numbers in unsolicited emails or calls."],
             ["Do not disclose personal information, PIN, password or other details that may compromise online banking security to anyone."],
             ["Do not use public or Internet cafe computers to access online banking or perform financial transactions."]])]],

        ["How should a computer be safeguarded for Standard Chartered Online Banking?",
         ["Standard Chartered advises the following computer safeguards:",
          B([["Install industry-recommended firewall and virus detection software to protect against hackers, virus attacks and malicious Trojan Horse programmes, and update them with security patches or newer versions regularly."],
             ["Make regular backups of critical data and consider encryption technology for highly sensitive data."],
             ["Disable the “File and Printer sharing for Microsoft Networks” feature to help prevent external parties gaining illegal control or access."],
             ["Delete and block junk or chain mails, and disable file sharing if it is not needed."],
             ["Do not install software or run programs of unknown origin, and do not open email attachments from strangers."],
             ["Educate everyone who uses the computer about basic security."]])]],

        ["What other security practices does Standard Chartered recommend for account holders?",
         ["Standard Chartered recommends the following additional practices:",
          B([["Check account and transaction history details regularly."],
             ["Update the Bank immediately when contact details change."],
             ["Report any unauthorised transactions on the account immediately."]])]],
      ]},
      { name: "Online Banking Security Guarantee", url: U.secGuarantee, qas: [
        ["What is the Standard Chartered Online Banking security guarantee?",
         [["When a client signs up for an online banking account, their money is automatically protected under this guarantee. In the event an unauthorised transaction is made from the online banking account, Standard Chartered will replace any loss of funds."]]],

        ["What conditions must a client meet for the Standard Chartered Online Banking security guarantee to apply?",
         ["The guarantee protects a client’s money where they have played their role in ensuring online security. The Bank needs to verify that the client:",
          B([["Has not given their security details to someone else, including their username and PIN, 2-Factor Authentication (2FA), Enhanced Second Factor Authentication, and Additional Login PIN (ALP) via SMS or security token."],
             ["Has followed the Bank’s recommended security practices, adhered to the terms and conditions, and followed the data protection and privacy policy applying to the account and online banking."],
             ["Does not use an account aggregation service, meaning a service provided by another company that allows all bank details to be viewed on a single website, or a similar service."],
             ["Monitors records of transactions regularly and reports discrepancies within 14 days from the date of the statement of account."],
             ["Offers full co-operation and provides all information needed to the Bank and relevant authorities."]])]],

        ["What are the limits of the Standard Chartered Online Banking security guarantee?",
         ["The following limits apply to the guarantee:",
          B([["The replacement of funds is limited to the amount illegally transferred from the Online Banking account. The guarantee does not cover any other losses, including indirect, consequential or special losses, damages, expenses, legal fees or loss of opportunity."],
             ["The guarantee does not apply where the transaction was carried out with the client’s knowledge and consent, or where the client has acted fraudulently, dishonestly or in a criminal manner, alone or with others."],
             ["Investigations begin the moment the Bank is informed that the account has been compromised."]])]],
      ]},
      { name: "Online Banking Security Features", url: U.secFeatures, qas: [
        ["What cryptography does Standard Chartered Online Banking use?",
         ["Standard Chartered Online Banking currently uses:",
          B([["SSL 3.0, RC4 with 128-bit encryption (high)."],
             ["RSA with 1024-bit exchange to ensure data (password) protection and data integrity protection."],
             ["This form of cryptography is among the strongest available in the industry."]])]],

        ["What access controls protect a Standard Chartered Online Banking account?",
         ["Access to Standard Chartered Online Banking requires a unique combination of Username and PIN. Additional measures protecting the account are:",
          B([["Limited repeated attempts to log in."],
             ["Automatic log out if there is no activity for a period of time, after which the client must log in again to continue with the service."]])]],

        ["What security infrastructure does Standard Chartered Online Banking use?",
         ["Standard Chartered’s security measures are consistent with international best practices and consist of:",
          B([["Packet filtering routers."],
             ["Firewalls."],
             ["Other security solutions aimed at safeguarding the system and protecting the interests of customers."]])]],

        ["How does Standard Chartered monitor and audit Online Banking security?",
         [["Standard Chartered has various security monitoring systems and processes in place that detect unauthorised activity. In addition, regular security reviews conducted by internal and external auditors ensure the surveillance systems are secure."]]],
      ]},
    ],
  },

  // ================================================== MANAGE YOUR FINANCES
  {
    name: "Manage Your Finances",
    subs: [
      { name: "Overview of Solutions", url: U.myfIndex, qas: [
        ["Which finance management solutions does Standard Chartered Singapore offer?",
         ["Standard Chartered Singapore offers the following solutions for managing card finances:",
          B([["Credit Card Funds Transfer: interest-free cash from as low as 1.80% processing fee (EIR from 4.86%)."],
             ["EasyPay: converts transactions into 0% instalments at a low service fee."],
             ["EasyBill: earns rewards on IRAS, Education, Insurance and Rental payments."],
             ["Temporary Credit Limit Increase: for emergencies, overseas travel expenses or a wedding banquet."],
             ["Instalment Payment Plan (IPP): splits eligible payments into instalments with 0% interest, though processing fees may apply."],
             ["Credit Card Credit Limit Review: increases the credit card limit instantly."]])]],
      ]},
      { name: "Credit Limit Review", url: U.clr, qas: [
        ["How is a Credit Card Credit Limit Review applied for at Standard Chartered Singapore?",
         ["A Credit Limit Review is applied for as follows:",
          B([["Step 1: Retrieve information via MyInfo or Online Banking."],
             ["Step 2: Verify personal details."],
             ["Step 3: Choose the preferred credit limit."],
             ["Step 4: Review the details."],
             ["Step 5: Submit the application."]])]],

        ["Who is eligible for a Standard Chartered Credit Card Credit Limit Review?",
         [["All principal cardholders of the Bank’s credit cards can apply for a Credit Card Credit Limit Review. Principal cardholders with supplementary cards must submit a manual form, as supplementary cardholders’ consents are required for a permanent credit limit increase. Credit limit review is not applicable for corporate cards, secured credit cards and Manhattan S$500 credit cards. Terms and conditions are published at ", ["sc.com/sg/terms-and-conditions/credit-limit-review", L.clrTnc], "."]]],

        ["What documents do Singapore Citizens and Permanent Residents need for a Standard Chartered Credit Limit Review?",
         ["For SingPass holders applying with MyInfo, no documents are required. For non-SingPass holders:",
          B([["Salaried employees: latest computerised payslip, or latest 6 months’ CPF Contribution History Statement, or latest Income Tax Notice of Assessment together with one of those two. To be considered for a higher credit limit, the latest Income Tax Notice of Assessment must be submitted in addition to the above."],
             ["Commission-based earners and self-employed: latest Income Tax Notice of Assessment and latest 3 months’ CPF Contribution History Statement."]])]],

        ["What documents do foreigners need for a Standard Chartered Credit Limit Review?",
         ["A foreign applicant must provide one of the following:",
          B([["Employment Pass with at least 6 months validity, together with the latest computerised payslip."],
             ["Company’s letter certifying employment and salary in Singapore dollar currency, dated within 3 months from the date of credit card application."],
             ["Latest Income Tax Notice of Assessment together with either of the two options above."]])]],
      ]},
      { name: "SC EasyBill", url: U.easybill, qas: [
        ["What is the SC EasyBill Payment Programme?",
         [["SC EasyBill allows Standard Chartered clients to earn 360° Rewards Points, cashback (as the case may be) or Bonus Interest (in the case of Bonus$aver credit cards) when they make eligible payments to education institutions, insurance premiums, rental or the Inland Revenue Authority of Singapore (IRAS) with their Bank credit cards."]]],

        ["Is there a fee for SC EasyBill?",
         [["Yes. There is a non-refundable processing fee of up to 1.9% applicable for each eligible SC EasyBill transaction. The fee is not a fixed amount and varies according to the SC EasyBill facility chosen on the Online Application. During the application, the fee is reflected once the type of eligible payment (IRAS, Rental, Education or Insurance) is selected and before the online application is submitted. The processing fee cannot be waived."]]],

        ["Is there a cap on SC EasyBill?",
         ["Yes. A Qualified Cardholder is limited to 1 SC EasyBill application per payment type every calendar month. There is also an amount cap for selected payment types:",
          B([["Rental Payments: SGD10,000."],
             ["Education Institution Payments: SGD30,000."]])]],

        ["How does SC EasyBill differ from the EasyPay Programme?",
         [["SC EasyBill is a rewards programme that allows clients to earn 360° Rewards Points, cashback (as the case may be) or Bonus Interest (in the case of Bonus$aver credit cards) on eligible payments to education institutions, insurance premiums, rental or IRAS using a Bank credit card. EasyPay is an instalment programme that converts eligible transactions of more than S$150 into instalments of 3 to 12 months, and converts outstanding balances of more than S$500 into instalments of 12 to 60 months. Clients can choose to use EasyPay to convert eligible SC EasyBill payments over S$150 into instalments."]]],

        ["Who can use SC EasyBill?",
         [["All principal cardholders of Standard Chartered credit cards can apply for SC EasyBill via SC Mobile or Online Banking."]]],

        ["Can overseas payments be made with SC EasyBill?",
         [["No. Eligible payments are limited only to the Singapore bank accounts listed on the SC EasyBill Online Application."]]],

        ["Does an active physical card need to be held to submit an SC EasyBill application?",
         [["Yes. An active physical card is required to submit an SC EasyBill application successfully. If only an active digital card is held, the SC EasyBill application will be rejected."]]],

        ["How is a client notified whether an SC EasyBill payment request is successful?",
         [["A confirmation, whether successful or unsuccessful, is sent via an electronic direct mailer once the SC EasyBill transaction is processed."]]],

        ["Can an SC EasyBill payment request be cancelled after submission?",
         [["No. SC EasyBill applications cannot be cancelled once submitted."]]],

        ["How long does an SC EasyBill payment take to reach the payee?",
         [["The eligible payment takes up to 7 business days to process, starting from the time of submission of the Online Application. For IRAS payments, the tax payable amount indicated on the payment request, along with the Tax Identification Number (TIN), is processed and credited directly to the tax account within 7 to 10 working days of the payment request, if the application is successful."]]],

        ["What happens if incorrect beneficiary details are provided on an SC EasyBill application?",
         [["The Bank is not responsible for any errors where incorrect beneficiary details were provided or where an incorrect beneficiary received the funds. The client must liaise directly with the recipient to retrieve the funds. Standard Chartered processes the payment request as per the information provided, so all beneficiary details must be verified by the client before submission."]]],

        ["Will Standard Chartered assist with SC EasyBill refunds for excess payment or an incorrect TIN?",
         [["The Bank is not able to initiate or assist with any refund proceedings. The client must reach out to the SC EasyBill beneficiary directly, for example IRAS, on this matter."]]],

        ["Where can the SC EasyBill fee be checked after a request is processed?",
         [["The SC EasyBill amount and fee can be checked via the monthly e-Statement or hardcopy statement."]]],

        ["Can recurring payments be made with SC EasyBill?",
         [["An option for recurring payments is currently not available. A new request must be submitted for each eligible payment."]]],

        ["When should payment be made to Standard Chartered for an SC EasyBill amount and fee?",
         [["Payment should always be made before the payment due date stated on the credit card statement. If the payment is not received in full, the relevant fees and charges are applied as per the credit card’s Terms & Conditions."]]],

        ["What should be done if an insurance company is not on the SC EasyBill list?",
         [["For insurance billers not on the list, all the details of the insurance biller should be filled in on the application. All SC EasyBill beneficiary details must be verified by the client, as failure to provide correct details may result in the SC EasyBill amount being credited into an incorrect bank account."]]],
      ]},
      { name: "EasyPay", url: U.easypay, qas: [
        ["What is Standard Chartered EasyPay?",
         [["EasyPay on Transaction allows retail purchases on a Standard Chartered credit card to be converted into interest-free instalments of 3 to 12 months with a one-time processing fee. EasyBill or Anytime Cash transactions can also be converted into instalments of 12 to 60 months at an attractive interest rate. Conversion is done via SC Mobile or Online Banking by selecting “Convert to Instalments” beside the eligible transaction."]]],

        ["What are the minimum amounts to apply for Standard Chartered EasyPay?",
         [["For EasyPay on Transaction applications, EasyPay is available for retail transactions above S$150. For EasyPay on Balance applications, it is available for balances above S$500. For EasyPay on Debit Card and Deposit Account, the minimum is S$500 for debit card retail and bill payments, or S$1,000 for a funds transfer account."]]],

        ["Who is eligible to apply for Standard Chartered EasyPay?",
         ["Eligibility for EasyPay is as follows:",
          B([["All principal and supplementary cardholders of the Bank’s credit cards."],
             ["The EasyPay request must be raised within 30 days of the transaction date or available outstanding balance on the credit card."],
             ["Applications are made via the SC Mobile app."],
             ["EasyPay on Debit Card and Deposit Account additionally requires the applicant to be an existing holder of a Standard Chartered Credit Card."]])]],

        ["Is there a fee for Standard Chartered EasyPay?",
         [["Yes. There is a non-refundable, one-time processing fee for each EasyPay transaction. During the application, the fee is reflected upon selection of the preferred choice of transaction and before the application is submitted."]]],

        ["Are there fees for terminating a Standard Chartered EasyPay plan early?",
         [["For transactions converted to instalments on retail transactions, there is no early redemption fee. For balances converted into instalments, there is a 3% early redemption fee."]]],

        ["What tenures are available on Standard Chartered EasyPay?",
         ["Tenures on EasyPay are as follows:",
          B([["EasyPay on Transaction for retail purchases: flexible tenures from 3 to 12 months at 0% p.a. interest, with a S$0 early redemption fee."],
             ["EasyPay on Balance: flexible tenures from 12 to 60 months, with attractive interest rates."],
             ["EasyPay on Transaction for debit card spends or bill payments: flexible tenures from 12 to 60 months, with attractive interest rates."]])]],

        ["How long does a Standard Chartered EasyPay application take to process?",
         [["An EasyPay application takes 3 to 5 working days to process. The outcome of the application, whether successful or unsuccessful, is communicated by email."]]],

        ["Can a Standard Chartered EasyPay application be cancelled after submission?",
         [["No. EasyPay applications cannot be cancelled once they are being processed."]]],

        ["Must payment still be made after submitting a Standard Chartered EasyPay application?",
         [["Yes. Regardless of whether the EasyPay application is approved or rejected, the cardholder must still make payment on the total outstanding balances and/or minimum payment due reflected on the credit card statement by the payment due date. Failing this, interest and/or late payment and finance charges may apply."]]],

        ["Which transactions are eligible for Standard Chartered EasyPay?",
         [["Retail transactions above S$150 are eligible for EasyPay. The option to apply is displayed alongside the eligible transactions on SC Mobile and Online Banking. Foreign currency transactions are included. The EasyPay request must be raised within 30 days of the transaction date on the credit card. The programme terms are published at ", ["sc.com/sg/terms-and-conditions/easypay-on-transaction", L.easypayTxnTnc], " and ", ["sc.com/sg/terms-and-conditions/easypay-on-retail-balances", L.easypayBalTnc], "."]]],
      ]},
      { name: "Instalment Payment Plan (IPP)", url: U.ipp, qas: [
        ["What is the Standard Chartered Instalment Payment Plan?",
         ["SC IPP converts transactions into interest-free monthly instalments when spending with a Standard Chartered Credit Card at participating merchants. Its benefits are:",
          B([["0% interest on the purchase split into monthly instalments."],
             ["Flexible payment options with instalment tenures of 3 months and more."],
             ["A minimum spend of S$500 at participating merchants, or S$100 for SC Visa credit card exclusive merchants."],
             ["Access to over 1,000 participating merchants."]])]],

        ["How are participating merchants under Standard Chartered SC IPP identified?",
         [["SC IPP is only applicable at participating merchants. The respective merchant should be asked whether SC IPP is available before making a purchase."]]],

        ["How is an existing Standard Chartered SC IPP cancelled, and are there early redemption charges?",
         [["The Bank must be notified in writing that the client has decided to repay the Instalment Purchase Price, as defined in Part G – 0% Interest Instalment Plan Product Terms of the Credit Card Terms, in full before the end of the selected tenure, and full repayment must be made before the end of that tenure. There are no charges for early redemption. The full outstanding amount is reflected in the next statement and must be settled in full before the statement due date, as payment after the due date incurs the relevant late and interest charges."]]],

        ["Can the tenure of a Standard Chartered SC IPP be changed?",
         [["No. Once an application for SC IPP is approved, the tenure of the SC IPP cannot be changed."]]],

        ["Are Rewards Points or cashback earned on Standard Chartered SC IPP transactions?",
         [["No. Since 2 May 2020, all SC IPP transactions are not eligible for cashback and 360° Rewards Points."]]],

        ["Is Standard Chartered SC IPP available on all credit cards?",
         [["SC IPP is not available on Standard Chartered corporate credit cards, Standard Chartered Platinum Access credit cards, or Standard Chartered cards with a credit limit of S$500."]]],
      ]},
      { name: "SC Visa Instalment Payment Plan", url: U.visaIpp, qas: [
        ["What is the SC Visa Instalment Payment Plan?",
         [["SC Visa Instalment Payment Plan (SC Visa IPP) allows a purchase of a minimum of S$100 at participating e-commerce merchants to be converted into interest-free monthly instalments with a Standard Chartered Visa Credit Card."]]],

        ["How is a purchase made using SC Visa IPP?",
         ["A purchase is made using SC Visa IPP as follows:",
          B([["Step 1: Look out for the “Instalments enabled by Visa” logo on the merchant’s webpage, and spend a minimum of S$100 on a Standard Chartered Visa Credit Card."],
             ["Step 2: Select the preferred instalment payment plan, with flexible tenures from 3 months onwards."],
             ["Step 3: Complete the payment."]])]],

        ["How does a client know if a card or purchase is ineligible for SC Visa IPP?",
         [["If the Standard Chartered Credit Card or the purchase is not eligible for SC Visa IPP, no instalment plans are shown at the checkout page."]]],

        ["What is the difference between SC IPP and SC Visa IPP?",
         [["SC Visa IPP is an extension of the SC IPP programme. SC Visa IPP offers interest-free instalment plans at Visa-exclusive e-commerce merchants when checking out with a Standard Chartered Visa Credit Card. The minimum purchase amount for SC Visa IPP is lower, from S$100, compared to SC IPP from S$500."]]],
      ]},
      { name: "Temporary Credit Limit Increase", url: U.tcli, qas: [
        ["What is a Standard Chartered Temporary Credit Limit Increase?",
         ["A Temporary Credit Limit Increase is a temporary increase of the credit limit on a credit card account, which can be used for the following purposes:",
          B([["Medical or hospitalisation expenses."],
             ["Funeral expenses."],
             ["Overseas travel expenses."],
             ["Wedding banquet expenses."]])]],

        ["Who can apply for a Standard Chartered Temporary Credit Limit Increase?",
         [["Application is for existing Standard Chartered principal credit cardholders only. Approval is instant, and the increased credit limit is assigned to the credit card accounts upon instant approval. Terms and conditions are published at ", ["sc.com/sg/terms-and-conditions/temporary-credit-limit-increase", L.tcliTnc], "."]]],

        ["How long is a Standard Chartered Temporary Credit Limit Increase valid?",
         [["A temporary credit limit increase is valid for the period approved by the Bank. The validity period starts from the date when the temporary credit limit is approved (the “Effective Date”) until the last day of the period the Bank has granted the temporary credit limit for (the “Expiry Date”). An option from 16 to 60 days of validity is available."]]],

        ["Are there fees for a Standard Chartered Temporary Credit Limit Increase?",
         [["There are no fees and charges for the Temporary Credit Limit Increase service. However, the utilised additional limit must be paid down upon expiry of the temporary credit limit, to avoid an overlimit fee of S$40, which is charged to the card account when the current balance exceeds the credit limit."]]],

        ["What information is required for a Temporary Credit Limit Increase for overseas travel?",
         ["The following information is required for overseas travel:",
          B([["Travel destination."],
             ["Intended travel booking date."],
             ["Departure and arrival dates."],
             ["Requested effective date, which must be within 7 days of intended usage."]])]],

        ["What information is required for a Temporary Credit Limit Increase for a wedding banquet?",
         ["The following information is required for a wedding banquet:",
          B([["Date of the banquet."],
             ["Name and contact details of the hotel or restaurant where the banquet will be held."],
             ["Name of the person in charge of the banquet at the hotel or restaurant."],
             ["Total number of tables and cost per table."],
             ["Requested effective date, which must be within 7 days of intended usage."]])]],

        ["What information is required for a Temporary Credit Limit Increase for medical, hospitalisation or funeral expenses?",
         ["The information required is:",
          B([["For medical or hospitalisation: the name of the hospital or clinic, the relationship with the patient (only eligible for family members, either immediate or non-immediate), and the validity period comprising the Effective Date and Expiry Date."],
             ["For funeral: the validity period comprising the Effective Date and Expiry Date."]])]],
      ]},
    ],
  },

  // ================================================== MANAGE YOUR PAYMENTS
  {
    name: "Manage Your Payments",
    subs: [
      { name: "Payment Solutions Overview", url: U.mypIndex, qas: [
        ["Which payment solutions does Standard Chartered Singapore offer?",
         ["Standard Chartered Singapore offers the following payment solutions:",
          B([["FAST Transfer, PayNow fund transfers, MEPS+ for same-day high-value SGD payments, and Direct Transfers to e-wallets of non-bank financial institutions."],
             ["Bill Payments, GIRO, eGIRO and EDP as a digital alternative to cheques."],
             ["Scan and Pay using QR codes, and Credit Card Payments."],
             ["Mobile wallets: Apple Pay, Google Pay and Samsung Pay."],
             ["Contactless and online card payments: Tap & go with SimplyGo, Tap & Go with Mastercard, Visa payWave, Click to pay with Visa, and Masterpass by Mastercard."]])]],
      ]},
      { name: "Apple Pay", url: U.applePay, qas: [
        ["What is Apple Pay on a Standard Chartered card?",
         [["Apple Pay allows payment in stores or within apps using cards already held, on the devices used every day. Card details are never shared when using Apple Pay and are not stored on the device at all, making Apple Pay on iPhone, Apple Watch or iPad a safer and more private way to pay."]]],

        ["How is a Standard Chartered card added to Apple Pay on an iPhone?",
         ["A Standard Chartered card is added to Apple Pay on an iPhone as follows:",
          B([["Step 1: Open the Wallet app."],
             ["Step 2: Tap the “+” sign in the upper right corner."],
             ["Step 3: Capture the card details with the camera or enter them manually."],
             ["Step 4: Make the Standard Chartered card the default card."]])]],

        ["How is a payment made with Apple Pay using a Standard Chartered card?",
         ["Payment with Apple Pay works as follows:",
          B([["iPhone in stores: hold the iPhone near a contactless reader with a thumb on Touch ID, or double-click the Home button when the iPhone is locked to access Wallet and make the purchase."],
             ["Apple Watch in stores: double-click the side button and hold the display of the Apple Watch up to the contactless reader. A gentle tap and beep confirms the payment information was sent."],
             ["iPhone or iPad within apps: select the Apple Pay payment option at checkout, review the details, and place a finger on Touch ID."]])]],

        ["Which Standard Chartered cards can be registered for Apple Pay?",
         [["All Standard Chartered retail credit and debit cards can be registered for Apple Pay."]]],

        ["Can Standard Chartered supplementary credit cards be registered on Apple Pay?",
         [["Yes. Supplementary credit cards can be registered. However, the activation OTP is sent to the principal cardholder’s mobile number registered with Standard Chartered. Subsequent transactions on the supplementary card are treated as transactions on the principal credit card. A supplementary cardholder cannot see their supplementary card in SC Mobile and therefore cannot add it to Apple Wallet through SC Mobile."]]],

        ["Can a Standard Chartered debit card be added to Apple Wallet through SC Mobile?",
         [["No. Only credit cards can be added to Apple Wallet through SC Mobile. Debit cards can only be added directly from Apple Wallet."]]],

        ["What happens to Apple Pay if a Standard Chartered card is lost and replaced?",
         [["If the original payment card is lost and a replacement card is received, the original payment card must be removed from Apple Pay and the new card re-registered."]]],

        ["What should be done if an Apple device holding a Standard Chartered card is lost?",
         [["The best approach is to log in to iCloud using the Apple ID, as iCloud may help locate the lost device. If a card has been added to Apple Pay, all added cards can be suspended using the Find My iPhone function, and permanently removed using iCloud settings. Alternatively, Standard Chartered can suspend or delete the Digital Account Number (Token) on the device if the last 4 digits of the token are known, which does not affect normal usage of the physical card."]]],

        ["Can Apple Pay be used with a Standard Chartered card in foreign countries?",
         [["Yes. Apple Pay can be used on most contactless terminals in foreign countries."]]],

        ["Is an internet connection needed to use Apple Pay with a Standard Chartered card?",
         [["An active internet connection is needed for card registration and activation. An active internet connection is not required to make purchases. Ideally the device should connect to the internet so that transaction history can be viewed on the device and software kept updated."]]],

        ["How is the default Apple Pay payment card determined for a Standard Chartered card?",
         [["If a card has been registered for the iTunes store, the default card for Apple Pay will be the same card registered for the iTunes store. The default card can be changed by opening the Apple Pay app, holding onto the desired default payment card and dragging it to the front."]]],

        ["Why does the card image in Apple Pay not match the physical Standard Chartered card?",
         ["The card displayed may not always exactly match the physical card. The key points to look for are:",
          B([["The bank, shown by the Standard Chartered Bank logo."],
             ["The card network, either Visa or Mastercard."],
             ["The last four digits of the physical card."]])]],

        ["How many devices can a Standard Chartered card be added to on Apple Pay?",
         [["A Standard Chartered credit card can only be added on 10 devices for Apple Pay."]]],

        ["Why does the “Add to Apple Wallet” button still appear after adding a Standard Chartered card through SC Mobile?",
         [["This happens when the iPhone is paired with an Apple Watch. If the credit card has already been added through SC Mobile, it can then only be added on the Apple Watch. The same applies in reverse: if the Apple Watch is paired with an iPhone and the card was added through SC Mobile, it can then only be added on the iPhone."]]],

        ["Why does an Apple Pay transaction show as successful but is declined by the merchant terminal?",
         [["The merchant terminal’s response is the most accurate response. The phone may need some time to refresh the status of the transaction. With a data connection, a successful or declined transaction is shown in the transaction history at the back of the card, accessed by clicking the “i” at the bottom of the card page, provided the feature is enabled."]]],

        ["How is a transaction made with Apple Pay on a Standard Chartered card cancelled or voided?",
         [["Cancelling a payment made with a phone is the same as with a physical card: return to the merchant the purchase was made from and cancel the payment. If asked for the last 4 digits of the card, the last 4 digits of the Token or Device Account Number should be used instead of the last 4 digits of the physical card. To locate the Device Account Number, launch Apple Pay and click “i” at the bottom right of the card."]]],

        ["Can a blocked Standard Chartered card be added to Apple Pay?",
         [["A blocked card cannot be added to the wallet until the block code is cleared. The token may show as deactivated or suspended, or as ready for wallet, depending on the block code."]]],

        ["Why must an Apple Pay token be self-activated on a Standard Chartered card?",
         [["For better client protection, tokens are best self-activated via OTP."]]],

        ["What should be done about the error “Could Not Connect to Apple Pay” when launching Apple Pay?",
         [["This is a phone setting issue where 4G data is not enabled for “Wallet”. To resolve it, go to Settings, then Mobile Data, and turn on mobile data for “Wallet”."]]],
      ]},
      { name: "Google Pay", url: U.googlePay, qas: [
        ["What does Google Pay offer Standard Chartered clients?",
         ["Google Pay allows users to link their eligible bank accounts and make payments to any individual or merchant registered with PayNow, and is available on both iOS and Android devices. Its benefits are:",
          B([["Sending money to anyone registered with PayNow, on both Android and iOS devices."],
             ["Cashback with eligible transactions, subject to terms and conditions."],
             ["Machine learning and 24/7 fraud detection to help keep the account safe."]])]],

        ["How is a Standard Chartered bank account linked to Google Pay?",
         ["A Standard Chartered bank account is linked to Google Pay as follows:",
          B([["Step 1: Unlock the Google Pay app, click on the profile at the top right of the home screen, and select “Bank account” under “Set up payment methods”."],
             ["Step 2: Select “Add bank account”, select “Standard Chartered”, click “continue” and follow the on-screen instructions to link the account."],
             ["Step 3: Once verified, the account is ready for transferring money."]])]],

        ["How is a Standard Chartered credit or debit card linked to Google Pay?",
         ["A card is linked to Google Pay as follows:",
          B([["Step 1: Tap on the profile photo from the home screen in the Google Pay app."],
             ["Step 2: Select “Payment methods”."],
             ["Step 3: Tap “Add a debit card or credit card”."],
             ["Step 4: Follow the instructions."]])]],

        ["What is the daily limit on Google Pay transactions for Standard Chartered clients?",
         [["The daily limit for Google Pay transfers is SGD 1,000. This is a shared limit for all PayNow third-party fund transfers done both via Standard Chartered Online Banking or SC Mobile and Google Pay."]]],

        ["How many devices can a Standard Chartered card be added to on Google Pay?",
         [["Standard Chartered credit and debit cards can only be added on 10 devices for Google Pay per card."]]],

        ["Can Standard Chartered card details be removed from Google Pay when changing devices?",
         [["Yes. Stored details can be deleted on Google Pay, and the records held by the network (Visa or Mastercard) and Standard Chartered are updated automatically."]]],

        ["Which everyday services can be paid for with Google Pay using a Standard Chartered card?",
         ["Google Pay supports the following with a Standard Chartered card:",
          B([["Sending money to anyone in the user’s contacts, even if they do not have Google Pay, and splitting bills within a group."],
             ["Scanning a PayNow SGQR to make purchases at businesses that do not accept credit or debit cards."],
             ["Ordering food and paying for takeaway from restaurant menus in the Google Pay app."],
             ["Buying and checking out movie tickets, reserving seats, and showing tickets by phone at the theatre."],
             ["Tap to pay wherever contactless payments are accepted, including buses and MRT islandwide. This feature is only available for Android devices with NFC capabilities running Android Lollipop 5.0 or higher."]])]],
      ]},
      { name: "Samsung Pay", url: U.samsungPay, qas: [
        ["What is Samsung Pay on a Standard Chartered card?",
         [["Samsung Pay is a secure and easy-to-use mobile payment service which can be used to make purchases with almost every retailer in Singapore. It enables the use of both credit and debit cards by leveraging Magnetic Secure Transmission (MST) and Near Field Communication (NFC) technology."]]],

        ["How is a payment made with Samsung Pay using a Standard Chartered card?",
         ["Payment with Samsung Pay works as follows:",
          B([["Step 1: Swipe up to launch Samsung Pay."],
             ["Step 2: Authenticate with a fingerprint or 4-digit PIN."],
             ["Step 3: Tap the phone over the terminal to pay."]])]],

        ["What are the payment limits for MST and NFC on Samsung Pay?",
         ["Samsung Pay uses two technologies with different limits:",
          B([["Magnetic Secure Transmission (MST) accepts payments of any value, subject to the cardholder’s credit limit with the bank."],
             ["Near Field Communication (NFC) accepts payments up to S$100."]])]],

        ["Which Standard Chartered cards can be registered on Samsung Pay?",
         [["All Standard Chartered credit and debit cards, with the exception of corporate cards, can be registered on Samsung Pay."]]],

        ["Which devices support Samsung Pay?",
         [["Samsung Pay initially launched on the Samsung Galaxy S7 edge 4G+, S7 4G+, S6 edge+ 4G+ and Note 5 4G+, with other newer flagship models following."]]],

        ["How does Samsung Pay protect Standard Chartered card details?",
         ["Samsung Pay protects card details as follows:",
          B([["The actual card number is never shared with the merchant and is not stored on the device."],
             ["Digital tokenisation means each transaction uses an encrypted digital token to replace personal payment information."],
             ["The Samsung KNOX security platform monitors malicious software and activities on Galaxy devices installed with Samsung Pay."]])]],

        ["How many devices can a Standard Chartered card be added to on Samsung Pay?",
         [["Standard Chartered credit and debit cards can only be added on 10 devices for Samsung Pay per card."]]],

        ["Can Standard Chartered card details be removed from Samsung Pay when changing devices?",
         [["Yes. Stored details can be deleted on Samsung Pay, and the records held by the network (Visa or Mastercard) and Standard Chartered are updated automatically."]]],
      ]},
    ],
  },

  // ============================================== MOBILE BANKING SERVICES
  {
    name: "Mobile Banking Services",
    subs: [
      { name: "Standard Chartered Mobile — Setup and Login", url: U.scMobile, qas: [
        ["How is SC Mobile downloaded, and what are the device requirements?",
         [["SC Mobile can be downloaded from the App Store or Google Play Store. For iPhone users, the app requires a device running iOS 14 or later, with a minimum of an iPhone 6s Plus. For Android users, the app requires a device running Android OS Version 9.0 or above."]]],

        ["How can a client register for SC Mobile banking access without Singpass?",
         [["There are other ways to register for SC Mobile besides Singpass. A Standard Chartered ATM, debit or credit card can be used. Alternatively, registration can be completed using a temporary username sent to the registered email, along with a password sent to the registered mobile number by SMS."]]],

        ["Who can register for SC Mobile banking access using Singpass, and what information is retrieved?",
         [["Singapore Citizens, Permanent Residents and eligible Singpass users with valid details can register using Singpass. NRIC, registered mobile number and email address are securely retrieved to pre-fill the registration. Singpass is recommended for a seamless and secure registration experience without the need to enter card details. If details in the Singpass account are not up to date or are no longer valid, they should be updated in Singpass before proceeding with registration."]]],

        ["How does a client obtain and use a Temporary User ID and Password for Standard Chartered Online Banking?",
         ["If a new product has recently been applied for, the Temporary User ID and Password are sent separately via email and/or SMS. They are used as follows:",
          B([["Go to Online Banking."],
             ["Scroll to the bottom of the page and select “Register” with Temporary User ID and SMS PIN."],
             ["Review the information provided and accept the Terms & Conditions."],
             ["Enter the Temporary User ID and Password, then click “Next” to set up personal login credentials."]])]],

        ["How is an SC Mobile username or password reset?",
         ["A username or password is reset as follows:",
          B([["Click “Forgot?” on the SC Mobile login screen."],
             ["Provide either an ATM or Debit Card number with PIN, or a Credit Card number with PIN. If the PIN is not available, full credit card details are required."],
             ["Ensure the mobile number is registered with the Bank, as a temporary password is sent by SMS."]])]],

        ["What is SC Mobile Key?",
         [["SC Mobile Key is a 6-digit PIN that allows a client to authenticate their mobile and online banking logins and transactions through their mobile device."]]],

        ["Why can a client not log in after registering for SC Mobile Key?",
         [["When registering for online banking for the first time, or registering for SC Mobile Key on a new device, there is a 12-hour cooling period. During this time certain banking transactions cannot be performed."]]],

        ["How is the SC Mobile Key PIN changed?",
         ["The SC Mobile Key PIN is changed as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Under “Setting & Configuration”, select “Password and Security Settings”."],
             ["Select “SC Mobile Key” and click “Change PIN”."]])]],

        ["Can SC Mobile be used on more than one device?",
         [["Yes, SC Mobile can be used on different devices, though verification is prompted on the primary device. The SC Mobile Key can only be registered on one device, so switching devices requires re-registering the SC Mobile Key on the new device."]]],

        ["What is Singpass Face Verification on SC Mobile?",
         [["Singpass Face Verification (SFV) is a secure identity authentication method that uses facial recognition technology. It allows customers to sign up for SC Mobile and verify their identity for SC Mobile Key registration. It is available to all Singapore Citizens and Singapore Permanent Residents who have an active Singpass account. SFV is only available on SC Mobile and not on the SC Online Banking website."]]],

        ["Is Singpass Face Verification required at every SC Mobile login?",
         [["No. Singpass Face Verification is only required when setting up SC Mobile Key for the first time or when registering on a new device."]]],

        ["What should be done if Singpass Face Verification on SC Mobile is unsuccessful?",
         ["If Singpass Face Verification was unsuccessful, the following steps may resolve the issue:",
          B([["Ensure the surroundings are well-lit with minimal distractions in the background."],
             ["Look directly at the camera and remain still while scanning is in progress."],
             ["Remove glasses, masks or hats that may obstruct the face."],
             ["Check that the camera is functioning properly."],
             ["Singpass Face Verification may be attempted up to 5 times, after which there is a 30-minute wait before trying again."]])]],

        ["How is SC Mobile kept up to date, and does updating affect login credentials?",
         ["Enabling “Auto App Update” keeps SC Mobile on the latest version:",
          B([["For iOS: go to Settings > App Store > Automatic Downloads and turn on “Updates”."],
             ["For Android: open the Google Play Store, go to Settings > Auto-Update Apps and select either “Over any network” or “Over Wi-Fi only”."],
             ["Updates do not affect login credentials, and the existing Username and Password continue to work after the update."],
             ["Settings remain unchanged after updating, including Face ID and biometrics, SC Mobile Key, Push Notification preferences and the registered PayNow account."]])]],

        ["How is the SC Mobile app version checked?",
         ["The app version is checked as follows:",
          B([["Tap the Services tab at the bottom right corner."],
             ["Select “About”."],
             ["The current app version is displayed at the bottom of the screen."]])]],

        ["Which languages does SC Mobile support?",
         [["SC Mobile currently supports English and Mandarin."]]],
      ]},
      { name: "Standard Chartered Mobile — Accounts and Cards", url: U.scMobile, qas: [
        ["How are personal particulars updated in SC Mobile?",
         [["To update personal particulars, tap the Services tab at the bottom right corner and select “Personal Details”."]]],

        ["How is the daily transfer limit amended in SC Mobile?",
         [["Daily transfer limits can be adjusted in SC Mobile by logging in, going to the “Pay & Transfer” tab and updating the limit under “Manage Daily Transfer Limit”."]]],

        ["Can the authorisation limit be adjusted on a Standard Chartered account?",
         [["Authorisation limits are not adjustable. To safeguard clients against fraudulent transfers, authorisation limits are set at S$5,000 by default for all transaction types. This default does not apply to clients who have personalised authorisation limits below S$5,000."]]],

        ["Is the authorisation limit the same as the maximum daily limit on a Standard Chartered account?",
         [["No. The authorisation limit is an enhanced security feature which requires users to enter their SC Mobile Key to proceed with any transaction equal to or above S$5,000. It is independent of the daily transfer limit, which restricts the amount that users can transact. As part of enhanced anti-scam measures, the Bank sets the default authorisation limit to S$5,000 for new clients and for existing clients with authorisation limits exceeding S$5,000. This change does not apply to clients who have already personalised their authorisation limits to below S$5,000."]]],

        ["Where are eStatements and eAdvices viewed in SC Mobile?",
         ["eStatements and eAdvices are viewed as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Select “View eStatements”."],
             ["Select “Latest Statement” or “Past Statement”."]])]],

        ["What is SGFinDex and how is it accessed in SC Mobile?",
         ["SGFinDex allows a client to view and track their accounts across different banks, CPF and HDB, giving a comprehensive overview of their financial status. It is accessed as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Select “Manage SGFinDex (Consolidated Financial View)”."]])]],

        ["Where is the secured mailbox found now that SC Mobile has been updated?",
         [["The secured mailbox is only available via Online Banking. Previous messages can be accessed through Online Banking in a web browser by hovering over the profile at the top right corner and clicking “Mailbox”."]]],

        ["What is Privacy Mode in SC Mobile?",
         [["Privacy Mode allows a client to hide the balance of their assets. It is enabled by clicking the “eye” icon beside the balance under “Total Assets”."]]],

        ["How is a credit card activated in SC Mobile?",
         ["A credit card is activated as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Under “Digital Services”, select “Credit Card Activation & PIN Setup”. Alternatively, select “View All” and select “Credit Card Activation & PIN Setup” under the Card Management tab."]])]],

        ["How is a debit card activated in SC Mobile?",
         ["A debit card is activated as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Under “Digital Services”, select “Debit/ATM Card Activation & PIN Setup”. Alternatively, select “View All” and select “Debit/ATM Card Activation & PIN Setup” under the Card Management tab."]])]],

        ["Where are credit card details found in SC Mobile?",
         [["Click on the credit card in question, tap the “eye” icon located on the card, and the credit card number and expiry date are displayed on the eCard. For CVV details, refer to the physical card."]]],

        ["Why are some credit card transactions missing from SC Mobile?",
         [["All posted and unposted transactions appear under “Card Activity”. However, some transactions may take time to be posted to the account. For example, if a merchant has approved the transaction but has not yet submitted the final charge, the transaction may not immediately show up in the transaction history."]]],

        ["How is a credit card temporarily locked or unlocked in SC Mobile?",
         ["A credit card is locked or unlocked as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Under “Digital Services”, select “View all”."],
             ["Under “Card Management”, select “Card Settings”."],
             ["Select the card to lock or unlock. Once locked, the credit card is not displayed on the Home page, and it can be unlocked anytime by following the same steps."]])]],

        ["How is a debit card temporarily locked or unlocked in SC Mobile?",
         ["A debit card is locked or unlocked as follows:",
          B([["Log in to SC Mobile."],
             ["On the Home page, select the account under “Deposits”."],
             ["Look under the “Cards” tab and select “Lock/Unlock Card”. Alternatively, select the account under “Deposits” on the Home page, select “Manage”, then under “Debit Cards Management” select “Card Temp Lock / Unlock”."]])]],

        ["How is a credit card fee waiver requested in SC Mobile?",
         ["A fee waiver is requested as follows:",
          B([["Log in to SC Mobile."],
             ["Tap the Services tab at the bottom right corner."],
             ["Under Digital Services, select “View All”."],
             ["Under “Help & Services”, select “Card Management”."],
             ["Select “Credit Card Fee Waiver” and choose the card the fee waiver is requested for."]])]],

        ["How is the payment due for the current month found in SC Mobile?",
         [["The payment due for the current month’s statement is the amount indicated next to the Due Date on the Home page. Alternatively, tap the credit card account, click “Details” and check the Statement Balance."]]],

        ["How are credit card rewards viewed or redeemed in SC Mobile?",
         [["Click on the credit card in question and select “Rewards” to view available points and redeem rewards."]]],

        ["Can a bank account be opened through SC Mobile?",
         ["Yes. An account is opened as follows:",
          B([["Log in to SC Mobile."],
             ["Select “Discover” at the bottom navigation bar."],
             ["Tap “Deposits Accounts” to view the available products and select the type of account to open."],
             ["Follow the instructions and submit the application."]])]],

        ["How is an additional currency added to a multi-currency account in SC Mobile?",
         [["From the Home page, tap “Add Other Currency” and select the currency to add."]]],

        ["Can a parent see a child’s First$aver account activity in SC Mobile?",
         ["Yes. Parents have full access to their child’s financial activities, including balance and transaction history. These details are viewed as follows:",
          B([["Log in to SC Mobile."],
             ["Look under “Family Banking” on the Home page."],
             ["Select the account to view the balance and transaction history."],
             ["A simplified version of SC Mobile is available for youths to manage their daily banking activities."]])]],
      ]},
      { name: "Standard Chartered Mobile — Payments and Investments", url: U.scMobile, qas: [
        ["How are credit card bills paid in SC Mobile?",
         ["Credit card bills are paid as follows:",
          B([["Tap “Pay & Transfer” at the bottom navigation bar."],
             ["Select “Pay All Credit Cards”. Payment can be made for a Standard Chartered credit card or another bank’s credit card."],
             ["Enter the details of the card to pay by tapping “Add Card”."]])]],

        ["How is a fund transfer repeated in SC Mobile?",
         ["A fund transfer is repeated as follows:",
          B([["Tap “Pay & Transfer” at the bottom navigation bar."],
             ["Select “Repeat Transfer”. The details from the previous transfer, including the from and to account and the amount, are automatically filled in."],
             ["Swipe to complete the transaction."]])]],

        ["How is a payee list managed in SC Mobile?",
         ["Payees are added, removed or edited as follows:",
          B([["Tap “Pay & Transfer” at the bottom navigation bar."],
             ["Select “Manage”."],
             ["Under “Payees”, select “Add & Manage Payees”."]])]],

        ["How is a billing organisation added or deleted in SC Mobile?",
         [["Tap “Pay & Transfer” at the bottom navigation bar, and under “Pay Bills” select “Add Biller” to add a new billing organisation. To delete a biller, log in via Online Banking."]]],

        ["How is a scheduled transaction created in SC Mobile?",
         ["A scheduled transaction is created as follows:",
          B([["Tap “Pay & Transfer” at the bottom navigation bar."],
             ["Select the account to transfer to under “Local Transfer” or “Between My Accounts”."],
             ["Click the “When” option to indicate the frequency and date for the transfer."]])]],

        ["How are scheduled transactions and GIRO arrangements viewed or cancelled in SC Mobile?",
         [["Tap “Pay & Transfer” at the bottom navigation bar and select “Scheduled” to view the list of GIRO and scheduled payment arrangements, where any scheduled transaction can be edited or deleted."]]],

        ["How is PayNow registration or de-registration done in SC Mobile?",
         ["PayNow registration is managed as follows:",
          B([["Tap the Services tab at the bottom right corner."],
             ["Under “Settings & Configuration”, click “Pay & Transfer Settings”."],
             ["Select “PayNow Registration”. This feature can also be accessed by clicking “Pay & Transfer” at the bottom navigation bar."]])]],

        ["How is a QR code paid using Scan & Pay in SC Mobile?",
         [["Tap “Pay & Transfer” at the bottom navigation bar and select the “Scan & Pay” option to access this function."]]],

        ["How is an investment profile created in SC Mobile?",
         ["An investment profile is created as follows:",
          B([["Tap “Invest” from the bottom navigation bar."],
             ["Select “Investment Profile”."],
             ["Select “Create My Investment Profile” to begin creating the profile."]])]],

        ["What is the Insights Hub in SC Mobile?",
         [["The Insights Hub is a feature within SC Mobile designed to provide a comprehensive overview of a client’s financial status in one place. It allows tracking and viewing of key financial insights, such as interest earned, savings, a summary of cashflow and card spending. Budgets are managed by logging in to SC Mobile and selecting “Insights” on the Home page to set limits."]]],

        ["Does the SC Mobile account balance include investment holdings?",
         [["No. The account balance does not include investment holdings. It only reflects the balances from Current, Savings and Fixed Deposit accounts, where applicable. Investment holdings are viewed under the “Investments” section."]]],

        ["How is the SC Online Trading Platform accessed via SC Mobile?",
         ["The trading platform is accessed as follows:",
          B([["Log in to SC Mobile."],
             ["Tap “Invest” at the bottom navigation bar."],
             ["Select “Equities”. Clients without a trading account are directed to an application form."],
             ["Tap “Trade Now” and search for a stock."],
             ["Select “Trade”, set the “Order Instruction” and click submit."]])]],
      ]},
      { name: "SC Mobile on Huawei AppGallery", url: U.huawei, qas: [
        ["Why is SC Mobile available on Huawei AppGallery?",
         [["As access to the Google Play Store is restricted in China, Standard Chartered has made the SC Mobile app available to download for Android phones through Huawei AppGallery. The app can be found by searching for “SCMobileSingapore” in the store."]]],

        ["Must a client in China who already has SC Mobile from Google Play download the Huawei AppGallery version?",
         [["No. There is no need to download the Huawei AppGallery version of SC Mobile. At any one time, a client should only be using one SC Mobile app. If SC Mobile was downloaded from the Google Play Store and works without issue, no action is required."]]],

        ["Why can a client in China not find the SC Mobile app in Huawei AppGallery?",
         [["Clients based in China need to change their Country/Region selection to “Singapore” in Huawei AppGallery. This is done in the profile “Me” tab, and requires a valid Huawei ID. The app can then be installed by searching for “SC Mobile Singapore”."]]],

        ["Which notifications are unavailable on the Huawei AppGallery version of SC Mobile?",
         [["Users located in China who downloaded SC Mobile Singapore from Huawei AppGallery cannot receive any push notifications or in-app notifications, and cannot enable the Inbox Notifications feature. This is because Google Play Services is required for in-app and push notification messages. Users continue to receive banking alerts via SMS on their registered mobile device."]]],

        ["How does SC Mobile Key work on the Huawei AppGallery version of SC Mobile?",
         [["SC Mobile Key continues to work on the mobile device registered with the digital token. For Online Banking and any other secondary device, users must generate an “Offline PIN” on the registered mobile device to authenticate any login or high-risk transactions. Alternatively, users can manually launch the SC Mobile app to authenticate their logins and transactions."]]],
      ]},
    ],
  },

  // ============================================== ONLINE BANKING SERVICES
  {
    name: "Online Banking Services",
    subs: [
      { name: "Online Banking Features", url: U.onlineBanking, qas: [
        ["What account information and services are available on Standard Chartered Online Banking?",
         ["Standard Chartered Online Banking provides the following account information and services:",
          B([["A summary of all accounts at a glance, and up to 360 days of transaction history for a deposit account."],
             ["Global link, which shows balances across selected Standard Chartered accounts worldwide. Accounts in up to 5 countries can be linked and accessed using the Online Banking Username and Password of the selected country."],
             ["Funds transfer locally and internationally between the client’s own Standard Chartered accounts, third-party Standard Chartered accounts, or any GIRO participating bank accounts."],
             ["Credit card funds transfer, which uses the credit card as a cash advance facility to transfer funds."],
             ["Pay-any-card, which pays Visa or Mastercard credit card bills from any bank in Singapore on a single platform."],
             ["Bill payment, a complimentary service allowing payment using a deposit account or credit cards to over 360 participating billing organisations."],
             ["eCashier’s Order, which orders a cheque online by providing the payee’s name and address."]])]],

        ["What is FAST on Standard Chartered Online Banking?",
         [["FAST moves money in 30 seconds via Standard Chartered Mobile or Online Banking, anytime and anywhere. FAST transfers are only available for Singapore dollar account transfers in Singapore between the 19 participating banks."]]],

        ["Which card and loan services are available on Standard Chartered Online Banking?",
         ["Standard Chartered Online Banking provides the following card and loan services:",
          B([["Credit card services: balance enquiry, details enquiry, transaction history, payment, applications and EasyPay application."],
             ["Loan services: balance enquiry, details enquiry, payment and applications."]])]],

        ["What personalisation options are available on Standard Chartered Online Banking?",
         ["Standard Chartered Online Banking offers the following personalisation options:",
          B([["Update personal particulars."],
             ["Create nicknames for accounts, cards and loans."],
             ["Sort accounts according to preference."],
             ["Lower the daily transfer limit for third-party fund transfers."],
             ["Customise authorisation limits to receive alerts."]])]],

        ["Which rates can be checked on Standard Chartered Online Banking?",
         ["The following rates can be checked on Online Banking:",
          B([["e$aver Deposit rate, Time Deposit rate, Saving Deposit rate, Cheque and Save Deposit rate, OneAccount Deposit rate and XtraSaver Deposit rate."],
             ["Foreign Exchange Deposit rate."],
             ["Lending rate."]])]],

        ["How does a client self-register for Standard Chartered Online Banking?",
         [["A client can register instantly for Online Banking using their ATM card, Debit Card, Credit Card or Phone Banking credentials."]]],
      ]},
      { name: "Online Bill Payment", url: U.billMerchants, qas: [
        ["How is a bill paid through Standard Chartered Online Banking or SC Mobile?",
         ["A bill is paid as follows:",
          B([["Step 1: Log in to Online Banking or the Mobile Banking App, and go to “Menu” > “Transfers & Payments”."],
             ["Step 2: In the “Transfers & Payments” screen, select “Pay Bills”."],
             ["Step 3: Select the biller to proceed with payment. If no billers have been set up, go back to the previous screen and select “Add and Manage Biller”."],
             ["Step 4: Enter the payment amount and select “Continue”."],
             ["Step 5: Select the payment date if required, set up a recurring bill payment if wanted, and select “Continue” to complete payment."]])]],

        ["How many billing organisations can be paid through Standard Chartered Online Banking?",
         [["Standard Chartered’s complimentary bill payment service allows payment using a deposit account or credit cards to over 360 participating billing organisations. The full list of participating billing organisations is published at ", ["sc.com/sg/bank-with-us/online-banking-services/online-bill-payment-merchants", L.billMerchants], "."]]],
      ]},
      { name: "eStatements and eAdvices", url: U.eStatements, qas: [
        ["What are Standard Chartered eStatements and eAdvices?",
         [["eStatements are an electronic version of an account statement, available on SC Mobile. eAdvices are an electronic version of a paper notice, emailed to the latest registered email address with the Bank or otherwise accessible via Online Banking. eAdvices is a bundled service with eStatements, and clients on eStatements are automatically enrolled for eAdvices."]]],

        ["How much does the Standard Chartered eStatements and eAdvices service cost?",
         [["There is no fee for using the eStatements and eAdvices service. Enrolment and usage of this service is free of charge."]]],

        ["Can Standard Chartered eStatements and eAdvices be accessed from anywhere in the world?",
         [["Yes. eStatements and eAdvices can be retrieved anytime and anywhere, as long as there is access to email or to Online Banking."]]],

        ["How often are Standard Chartered eStatements updated?",
         [["eStatements are updated on a monthly basis. Recent transactions can be viewed in “Account History” via Online Banking."]]],

        ["Are Standard Chartered eStatements delivered by email?",
         [["eStatements can be requested by email when updating the subscription preference. When an eStatement or eAdvice is ready, an email notification together with the password-protected eStatement or eAdvice is sent to the last updated email address held by the Bank. Real-time alerts are also sent when eStatements are ready."]]],

        ["How are Standard Chartered eStatements viewed or downloaded on SC Mobile?",
         ["eStatements are viewed or downloaded as follows:",
          B([["Step 1: Go to “View eStatements” in the Services tab located at the bottom right corner."],
             ["Step 2: Choose “Latest Statement” or “Past Statement” and click the download button beside the statement to view."],
             ["Step 3: The statement can be viewed instantly."]])]],

        ["How is an eStatement password changed on Standard Chartered Online Banking?",
         ["The eStatement password is changed as follows:",
          B([["Step 1: Select “eStatement/eAdvice” located above the account summary."],
             ["Step 2: Click “Settings”."],
             ["Step 3: Click “Change”, enter the new password and confirm to effect the change."]])]],

        ["What fee applies to Standard Chartered paper statements?",
         [["With effect from 1 January 2021, a S$2 Paper Statement Fee applies to clients who opt to receive paper statements. This fee is not applicable to clients below 18 years old or above 65 years old."]]],

        ["How is a registered email address or mobile number updated for Standard Chartered eStatements?",
         [["If an email address or mobile number has recently changed, particulars can be updated via Online Banking or the SC Mobile app under Help and Services."]]],
      ]},
      { name: "Alerts and Notifications", url: U.alerts, qas: [
        ["What are Standard Chartered banking alerts?",
         [["Banking alerts are notifications about transactional activities on a bank account. By default, these alerts are delivered by email and/or push notification. Alert preferences can be customised via Online Banking."]]],

        ["What are the benefits of Standard Chartered banking alerts?",
         [["Banking alerts allow a client to enjoy timely updates on their account activities and act quickly in the event of fraudulent transactions."]]],

        ["How is the Standard Chartered Alerts & Notifications feature enabled?",
         ["The feature is registered for via Online Banking as follows:",
          B([["Step 1: Log in to Online Banking with the username and password."],
             ["Step 2: Select “Alerts & Notifications” in the dropdown menu shown when hovering over the client’s name."],
             ["Step 3: Click “Manage” to continue."],
             ["Step 4: Navigate through the page to manage the alerts."]])]],

        ["Which activities can Standard Chartered Alerts & Notifications cover?",
         [["A client can choose to receive notifications on specific transactions, ATM withdrawals or regular balance updates on their accounts, reminders of due payments, and spending alerts for Standard Chartered Credit Cards. A preferred threshold can be chosen to trigger these alerts."]]],

        ["Can Standard Chartered Alerts & Notifications be managed in the SC Mobile app?",
         [["No. This feature can only be enabled, disabled or customised via Online Banking. If registered for Inbox Notifications, alerts may be converted to push notifications."]]],
      ]},
    ],
  },

  // ======================================================= STANDALONE PAGES
  {
    name: "Fees, Charges and Other Services",
    subs: [
      { name: "Understanding Fees and Charges", url: U.fees, qas: [
        ["What is a Standard Chartered Late Payment Charge?",
         [["A Late Payment Charge is an extra amount charged for payment on a credit card that was made after the due date. A charge of S$100 applies when the minimum payment due, as indicated on the statement, was not received by the due date. This S$100 charge is a fixed fee levied on the credit card account, irrespective of the amount outstanding in the account. This Late Payment Charge also applies to Credit Card Funds Transfer and CashOne instalment loan accounts."]]],

        ["What is a Standard Chartered credit card Annual Fee?",
         [["An Annual Fee is charged every year as a subscription for a credit card, so that cardholders can enjoy privileges and promotions for maintaining the credit card account. It is charged up-front yearly when the credit card fee date is due. The respective amounts charged are published in the ", ["Standard Chartered Singapore Pricing Guide", L.pricingGuide], "."]]],

        ["What is a Standard Chartered Finance or Interest Charge?",
         [["Any unpaid amount on a credit card is rolled over to the next statement and charged interest. Finance and Interest Charges are computed at the Effective Interest Rate (EIR) of 27.90% per annum, calculated daily from the date of the transaction until the balance owing is paid in full. No Finance or Interest Charges are imposed on any fees billed to the card account, such as annual fee, late payment charge and any other fees charged by the Bank."]]],

        ["How is a Standard Chartered credit card fee waiver requested?",
         ["A waiver request is submitted through SC Mobile or Online Banking as follows:",
          B([["Step 1: Log in to SC Mobile or Online Banking."],
             ["Step 2: Select “Help & Services”."],
             ["Step 3: Select “Card Management”."],
             ["Step 4: Select “Credit Card Fee Waiver”."],
             ["Step 5: Select the eligible card for the fee waiver."],
             ["The outcome of the request is subject to the Bank’s approval."]])]],

        ["How can a client be reminded of a Standard Chartered credit card payment due date?",
         ["The following options help a cardholder pay on time:",
          B([["Set up payment due alerts to give a reminder of the payment due date."],
             ["Subscribe to eStatements."],
             ["Apply for GIRO payment for the credit card to set up recurring payments."]])]],

        ["When are payments to a Standard Chartered credit card effected?",
         ["Payment crediting times are as follows:",
          B([["Payment via SC Mobile (Pay SC Credit Card): immediate crediting."],
             ["Payment via SC Online Banking (Pay SC Credit Card): immediate crediting."],
             ["Payment via Interbank: within 3 working days."],
             ["Payment via AXS: within 2 working days."],
             ["Payment by GIRO: deducted from the designated Singapore bank account two working days before the payment due date."],
             ["Payment by cheque: within 5 working days."]])]],
      ]},
      { name: "Understanding Late Payment Charge", url: U.latePayment, qas: [
        ["How many Standard Chartered late payment charge waivers can a client receive?",
         [["As a gesture of goodwill, the Bank offers clients one late payment charge waiver across all eligible accounts with the Bank within a period of 12 rolling months. If a waiver has been granted in the past 12 rolling months, the Bank will not grant any further waivers nor consider any appeals. For example, if a client holds Credit Card A and Credit Card B, and Credit Card B incurred a S$100 late payment charge in August 2022, but the client already received a waiver on Credit Card A in September 2021, the client is not eligible for a further waiver and is expected to pay the S$100 charge on Credit Card B."]]],

        ["Why was a Standard Chartered late payment charge waiver request declined?",
         [["The waiver request is reviewed against the client’s repayment history together with any past waivers provided. The Bank offers all clients one late payment charge waiver across all their eligible accounts, where the account is valid and in good standing and not delinquent for 30 days or more, within a period of 12 months. If a waiver has been received in the past 12 months, no further waivers can be granted. The outstanding bill should be paid promptly to avoid further charges."]]],

        ["What is the maximum Standard Chartered late payment charge waiver available?",
         [["The one late payment charge waiver given within the period of 12 months is based only on the latest charge found on the current or unbilled statement."]]],

        ["How is a Standard Chartered late payment charge avoided?",
         [["A late payment charge is avoided by ensuring payments are credited to the account by the payment due date indicated in the statement. Payment is effective only when cleared funds are received in the Standard Chartered account."]]],

        ["What happens if a Standard Chartered credit card payment is not made on time?",
         ["If the minimum payment due is not received by the payment due date, the following may occur:",
          B([["The credit card may be suspended."],
             ["Being consistently late for credit card payments may affect the client’s credit rating and their ability to obtain new unsecured credit facilities."]])]],

        ["How are Standard Chartered payment due date reminders set up?",
         [["A client may register for payment alerts via Standard Chartered Online Banking. Registration for alerts, which are delivered by email, is only available through Online Banking. After registering for payment alerts, the SC Mobile Banking App should be checked to confirm that notifications have been enabled to receive such alerts on the app. Subscribing to eStatements is also recommended."]]],

        ["How is GIRO set up for Standard Chartered credit card or personal loan payments?",
         [["GIRO arranges for Credit Card or Personal Loan payments to be debited automatically each month from a designated bank account. An Interbank GIRO form must be downloaded and submitted. If the debiting account is from Standard Chartered Bank (Singapore) Limited, this may take up to 2 weeks for processing. For all other banks, it may take 4 to 6 weeks."]]],

        ["Does a pending merchant refund prevent a Standard Chartered late payment charge?",
         [["No. Merchant refunds are not considered as payments to a card account. Payments must be made to the account by the payment due date indicated in the statement to avoid any late payment charges."]]],

        ["Can a client appeal a Standard Chartered late payment charge in a difficult situation?",
         [["Standard Chartered recognises that challenging situations may arise, such as major medical emergencies or other unfortunate situations. An appeal can be made via the secured mailbox option in Standard Chartered Online Banking or the Mobile App, and complete details must be provided for a review. The outcome of the request is subject to the Bank’s approval."]]],
      ]},
      { name: "SC Remit", url: U.remittance, qas: [
        ["What does SC Remit offer Standard Chartered Singapore clients?",
         [["SC Remit offers S$0 Transfer Charges on online remittances, and 0% FX Cost on online remittances in INR to Standard Chartered Bank, India accounts. Terms and conditions apply."]]],

        ["Which markets and currency pairs are eligible for S$0 SC Remit transfer charges?",
         ["The following payee markets are eligible, with the send currency shown first and the payment currency the payee receives shown second:",
          B([["United States of America: SGD or USD to USD. United Kingdom: SGD, USD or GBP to GBP. Jersey: SGD or USD to GBP."],
             ["India: SGD or USD to INR. Pakistan: SGD or USD to PKR. China: SGD or USD to CNY. Hong Kong: SGD, USD or HKD to HKD. Japan: SGD, USD or JPY to JPY."],
             ["Australia: SGD, USD or AUD to AUD. Canada: SGD, USD or CAD to CAD."],
             ["Indonesia: SGD or USD to IDR. Malaysia: SGD or USD to MYR. Philippines: SGD or USD to PHP. Thailand: SGD or USD to THB. Vietnam: SGD or USD to VND."],
             ["All SEPA countries: SGD, USD or EUR to EUR. United Arab Emirates: SGD or USD to AED."]])]],

        ["How is a remittance made using SC Remit?",
         ["A remittance is made as follows:",
          B([["Step 1: Log in to the SC Mobile App."],
             ["Step 2: Tap “SC Remit” on the home screen."],
             ["Step 3: Choose the account to transfer funds from and the payee account to transfer funds to."],
             ["Step 4: Once the payee details are confirmed, the rate is guaranteed for 2 minutes."],
             ["Step 5: Verify the transaction details and tap “Confirm” to submit the transaction."]])]],

        ["What charges does Standard Chartered waive on a Qualifying Remittance Transaction?",
         [["Under the SC Remit Promotion, all transfer charges and expenses imposed by the Bank for a Qualifying Remittance Transaction are waived. The Bank is not responsible for any fees and charges levied by the beneficiary bank in connection with the transaction. From 9 November 2020, the Bank also does not impose any additional FX costs for a Qualifying Remittance Transaction in INR as payment currency made to accounts maintained with Standard Chartered Bank, India."]]],
      ]},
      { name: "Money Lock", url: U.moneyLock, qas: [
        ["What is Standard Chartered Money Lock?",
         [["Standard Chartered Money Lock is an anti-scam security feature on SC Mobile that helps protect funds from scams and unauthorised transfers or withdrawals. Money from current or savings accounts that are not needed for everyday access can be locked, and the locked funds continue to earn interest at current rates."]]],

        ["Why should a client use Standard Chartered Money Lock?",
         [["With scams on the rise and tactics becoming more sophisticated, Standard Chartered Money Lock provides an additional layer of protection against scams. In the unlikely event that a scammer gains unauthorised access to an account digitally, the locked funds remain safe."]]],

        ["How are funds locked using Standard Chartered Money Lock?",
         ["Funds are locked as follows:",
          B([["Log in to SC Mobile."],
             ["Select the current or savings account."],
             ["Select the Money Lock option."],
             ["Indicate the amount to lock and submit the request."],
             ["A confirmation notification is received once the funds are successfully locked. Money is locked on the same day the application is submitted."]])]],

        ["What transactions are blocked by Standard Chartered Money Lock?",
         ["Locked funds are protected from, and cannot be used for:",
          B([["New or existing payment arrangements, such as GIRO, standing instructions or future-dated transfers."],
             ["Transfers to another Standard Chartered account, or to another bank’s account including local and overseas transfers."],
             ["Bill or credit card payments."],
             ["Time deposit placements, wealth and investment product purchases, and insurance purchases or payments."],
             ["Loan or tax repayments."],
             ["ATM withdrawals or transfers."],
             ["Locked funds also cannot be used for account fees, so a sufficient available balance should be kept to avoid late fees and charges."]])]],

        ["How are Standard Chartered Money Lock funds unlocked?",
         [["Locked funds can only be released after identity is verified in person at any Standard Chartered Singapore branch. This ensures that scammers cannot withdraw funds in the unlikely event they gain unauthorised access to the account. Locked funds are released in full on the same day."]]],

        ["Can part of the funds locked under Standard Chartered Money Lock be released?",
         [["No. Locked funds cannot be partially released. The entire locked amount must first be released, after which a request to lock a new amount can be submitted via SC Mobile."]]],

        ["Does a new account need to be opened to use Standard Chartered Money Lock?",
         [["No. A new account does not need to be opened. Funds in an existing current or savings account can be locked with SC Mobile."]]],

        ["Is it mandatory to lock funds with Standard Chartered Money Lock?",
         [["No, it is not mandatory to lock funds. However, Standard Chartered encourages the use of Money Lock as an additional layer of protection against scammers. Funds are safe without Money Lock, as the Bank uses a range of security features and measures to protect clients."]]],

        ["Why is Money Lock not visible in the SC Mobile app?",
         [["The app may not be on the latest version of SC Mobile. For iOS users, the app should be updated to version 9.13.5 via the App Store. For Android users, the latest app was released progressively from July 2024."]]],

        ["How does a client based outside Singapore use Standard Chartered Money Lock?",
         [["Money Lock is available on SC Mobile, so funds can be locked from anywhere by logging in, selecting the preferred account, indicating the lock currency and amount, and submitting the request. However, requests to release locked funds can only be submitted in person at any Standard Chartered Singapore branch."]]],
      ]},
      { name: "Digital Banking", url: U.digitalBanking, qas: [
        ["Which everyday banking tasks can be done digitally with Standard Chartered Singapore?",
         ["Standard Chartered Singapore supports the following digital banking activities:",
          B([["Online Banking, which carries out local or overseas transactions and bill payments, including a complimentary bill payment service to over 360 participating billing organisations using a deposit account or credit cards."],
             ["Self-service requests via SC Mobile, such as requesting a new cheque book or a credit balance refund from a credit card."],
             ["SC Remit, which sends money at S$0 Transfer Charges on online remittances."],
             ["PayNow, which sends and receives money securely and almost instantly using a mobile number or NRIC."]])]],

        ["Which self-service requests can be raised digitally with Standard Chartered Singapore?",
         ["The following service requests can be raised through Standard Chartered digital channels:",
          B([["Credit Balance Refund, for an overpaid credit card bill, and Cheque Book Request."],
             ["Report Lost/Stolen Card, which blocks the card and sends a replacement, and Card Replacement for a damaged card."],
             ["Update Profile Details, for a changed address or mobile number."],
             ["Card Activation & PIN Set for a newly received ATM, debit or credit card, and Card PIN Change for a forgotten PIN."],
             ["Closure of CashOne or Credit Card Instalment Loan, and Credit Card Cancellation."],
             ["Status Enquiry, to track the status of a service request."]])]],
      ]},
      { name: "Stay Home, Spin and Win 2020 Promotion", url: U.stayHome, qas: [
        ["What was the Standard Chartered Stay Home, Spin and Win 2020 Promotion?",
         [["The Stay Home, Spin and Win 2020 Promotion allowed clients to grow a digital score for a chance to spin and win up to S$100 worth of GrabFood vouchers fortnightly when they performed a Qualifying Activity. This was a 2020 promotion and its promotion period has ended. Terms and conditions applied."]]],

        ["How did the Standard Chartered Stay Home, Spin and Win 2020 Promotion work?",
         ["The 2020 promotion, whose promotion period has ended, worked as follows:",
          B([["Step 1: Register for the promotion by sending an SMS to 77222 in the format “SC 8-digit Mobile Number”, for example “SC 12345678”."],
             ["Step 2: After receiving an acknowledgement SMS confirming successful registration, perform Qualifying Activities every two weeks to increase the Digital Score. Accumulating 20 Digital Score points or more during the promotion period achieved Gold tier and a chance to win GrabFood vouchers of higher value."],
             ["Step 3: Eligible participants received a Winning Transaction SMS fortnightly showing their Digital Score, with a link to play the Stay Home, Spin and Win Game. Only one Winning Transaction SMS was issued every 14 days, with the Digital Score updated as of 11:59pm on the Sunday of each week of the promotion period."],
             ["Step 4: An email was sent with the redemption promo code for the relevant GrabFood voucher."],
             ["Step 5: Increasing the Digital Score by performing more Qualifying Activities allowed another Winning Transaction SMS in following weeks. Digital Score accumulated week on week."]])]],
      ]},
      { name: "Here for You", url: U.hereForYou, qas: [
        ["What support measures did Standard Chartered Singapore publish on its Here for You page?",
         ["The Here for You page collects Standard Chartered Singapore’s COVID-19 era support measures, dated to 2020. The measures listed are:",
          B([["Safe re-opening of branches, listing branches open to serve clients from 2 June 2020."],
             ["Relief measures offering support to existing Business Banking clients."],
             ["Updated Contact Centre services to prioritise urgent requests."],
             ["Notice of delays in mail services and processing of applications, as international mail services to certain countries were either unavailable or disrupted until further notice."],
             ["Updates on life insurance, where the Bank’s insurance partners introduced several industry-led measures aimed at mitigating the challenges presented by COVID-19."]])]],
      ]},
    ],
  },
];

const { doc, count } = buildDocument({
  title: "Standard Chartered Singapore Bank With Us FAQ",
  docTitle: "SCB Bank_Bank With Us FAQ",
  description: "Standard Chartered Singapore Bank With Us FAQ — RAG-ready",
  categories: CATEGORIES,
});

console.log("total Q&A:", count);
console.log("categories:", CATEGORIES.length, "subcategories:", CATEGORIES.reduce((n, c) => n + c.subs.length, 0));
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("SCB_Bank_Bank_With_Us_FAQ.docx", buf);
  console.log("written", buf.length);
});
