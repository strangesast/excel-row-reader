const fs = require('fs/promises');

const mod = require('./index');

const HEADER = [
  'Autofill Ind',
  'AWP',
  'Bil Patient Pay Amt',
  'Birth Dte',
  'BPL',
  'Cap Applied Amt',
  'Carrier',
  'Carrier Name',
  'Channel',
  'Claim Count',
  'Client Elig Membership ID',
  'Client Membership ID',
  'COB Payment',
  'Coinsurance Amt',
  'Compound Ind',
  'Contract',
  'Copay',
  'Copay BG Difference Paid',
  'Copay Cost Tier Seq',
  'Copayment Amt',
  'Cost Basis Cde',
  'Cost Basis Dsc',
  'Date of Service',
  'DAW Code',
  'Days Supply',
  'DEA Code',
  'Deductible',
  'Dep SSN',
  'Dispensing Fee',
  'Drug Name',
  'Drug Type',
  'Effective Date',
  'End Date',
  'External Group',
  'Fill QTY',
  'First Name',
  'Formulary Ind',
  'GCN',
  'Gender',
  'Generic Mfr Quantity Cde',
  'Generic Name',
  'Gross Cost',
  'Group',
  'Group Name',
  'Health Plan Funded Assist Amt',
  'Incentive Fee',
  'Ingredient Cost',
  'Invoice Date',
  'Last Name',
  'Locator',
  'Mail Service Rx Nbr',
  'Maintenance Drug',
  'Med B Ind',
  'Med D Ind',
  'Member ID',
  'Most Common Indication',
  'NDC',
  'Net Plan Cost',
  'NPI Nbr',
  'OOP Applied Amt',
  'Over Benefit Limit Amt',
  'Pass Through Txt',
  'Patient Brand Selection Amt',
  'Patient Coverage Gap Amt',
  'Patient ID',
  'Patient Level Auth Ind',
  'Patient Network Selection Amt',
  'Patient Non-Pref Brand Amt',
  'Patient Non-Pref Formulary Amt',
  'Patient Sales Tax Amt',
  'Person Nbr',
  'Pharmacy Claim ID',
  'Pharmacy Name',
  'Pharmacy Override Description',
  'Pharmacy RX No',
  'Phcy Override Cde',
  'PLA Auth Nbr',
  'Prescriber First Name',
  'Prescriber Last Name',
  'Prescriber NPI Nbr',
  'Prior Auth Ind',
  'Prior Auth Nbr',
  'Prior Auth Type Cde',
  'Processor Fee Amt',
  'Relshp Cde',
  'RRA Penalty Applied',
  'Rx Refill Nbr',
  'Sales Tax',
  'Specialty Drug',
  'STC',
  'STC Dsc',
  'Telecomm Version Cde',
  'Total Patient Cost',
  'Transaction Type Cde',
  'U&C Cost',
  'ZBL Excess Copay Paid',
];

(async () => {
  const buf = await fs.readFile('lookback-claims.xlsb');

  console.time('world');
  const result = mod.world(buf, 'CLAIM DETAIL', HEADER);
  console.timeEnd('world');

  console.time('map');
  const records = [];
  for (let i = 0; i < result.length; i++) {
    const record = {};
    for (let j = 0; j < HEADER.length; j++) {
      record[HEADER[j]] = result[i][j];
    }
    records.push(record);
  }
  await fs.writeFile('a.json', Buffer.from(JSON.stringify(records[0], null, 2)));
  // console.log(records[0]);
  console.log(records.length);
  console.timeEnd('map');
})();
