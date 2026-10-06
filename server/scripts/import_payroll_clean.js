const mongoose = require('mongoose');
const dotenv = require('dotenv');
const exceljs = require('exceljs');
const path = require('path');
const bcrypt = require('bcryptjs');

dotenv.config({ path: path.join(__dirname, '..', '.env') });

const Taxonomy = require('../models/Taxonomy');
const ProcessAnalysis = require('../models/ProcessAnalysis');
const WDTSubmission = require('../models/WDTSubmission');
const Fitment = require('../models/Fitment');
const User = require('../models/User');

async function run() {
  const uri = process.env.MONGODB_URI;
  if (!uri || !uri.includes('optinxt_payroll_db')) {
    console.error('ERROR: Database URI must contain optinxt_payroll_db. Aborting.');
    process.exit(1);
  }
  
  console.log('Connecting to', uri);
  await mongoose.connect(uri);
  
  console.log('Clearing old business data...');
  await Taxonomy.deleteMany({});
  await ProcessAnalysis.deleteMany({});
  await WDTSubmission.deleteMany({});
  await Fitment.deleteMany({});
  
  console.log('Parsing Excel...');
  const filePath = path.join('C:', 'Users', 'Swati', 'OneDrive', 'Desktop', 'project', 'Payroll Listing_01.10.2026.xlsx');
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];
  
  let pCount = 0;
  
  for (let rowNumber = 6; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    const scope = row.getCell(3).text?.trim();
    const subProcess = row.getCell(4).text?.trim();
    const activity = row.getCell(5).text?.trim();
    
    if (!activity) continue;
    
    const tax = new Taxonomy({
      department: 'HR',
      processScope: scope,
      subProcess: subProcess,
      activity: activity,
      systemId: 'PAYROLL-' + (pCount + 1),
      description: activity
    });
    await tax.save();
    
    const analysis = new ProcessAnalysis({
      taxonomyId: tax._id,
      criteria: '-',
      score: 0,
      fte: 0,
      consolidated: false,
      notes: ''
    });
    await analysis.save();
    
    pCount++;
  }
  
  console.log('Imported ' + pCount + ' activities.');
  
  const adminCount = await User.countDocuments();
  if(adminCount === 0) {
    console.log('Seeding default users...');
    const hp = await bcrypt.hash('Admin@123', 10);
    await User.create([
      { name: 'Admin', email: 'admin@bper.com', password: hp, role: 'admin', employeeId: 'A01', designation: 'Administrator', department: 'Corporate' },
      { name: 'Manager', email: 'manager@bper.com', password: hp, role: 'manager', employeeId: 'M01', designation: 'Manager', department: 'Operations' },
      { name: 'Employee', email: 'employee@bper.com', password: hp, role: 'employee', employeeId: 'E01', designation: 'Analyst', department: 'Operations' }
    ]);
  }
  
  console.log('Done.');
  process.exit(0);
}
run();
